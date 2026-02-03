# =============================================================================
# 剪贴板监控脚本 - 带FTP上传功能（修复超时问题版）
# 功能：监控剪贴板变化 → 保存到本地txt → 异步上传到FTP服务器
# 优化：FTP连接优化，支持主动/被动模式自动切换
# =============================================================================

# -----------------------------------------------------------------------------
# 配置区域 - 请根据需要修改这里的参数
# -----------------------------------------------------------------------------
$WorkDir = "C:\Users\user\Documents\copy"              # 本地工作目录
$OutputFile = "clipboard_log.txt"               # 输出文件名
$Interval = 1                                    # 检查间隔（秒）
$FtpTimeout = 15                                 # FTP上传超时时间（秒）- 增加到15秒

# FTP服务器配置
$FtpServer = "ftp://xxx.com:21"
$FtpUsername = "user"
$FtpPassword = "xxxxxxxxx"

# 启动选项
$ClearOldLog = $false  # true=每次启动清空旧日志; false=追加到旧日志

# FTP模式优先级（修改这里可以调整连接策略）
$PreferPassiveMode = $false  # false=优先主动模式; true=优先被动模式

# -----------------------------------------------------------------------------
# 全局变量：用于异步上传管理
# -----------------------------------------------------------------------------
$script:UploadJob = $null          # 当前上传任务
$script:PendingUpload = $false     # 是否有待上传的文件
$script:LastUploadTime = Get-Date  # 上次上传时间
$script:DebugMode = $true          # 调试模式

# -----------------------------------------------------------------------------
# 函数定义：上传文件到FTP（带重试和模式切换）
# -----------------------------------------------------------------------------
function Upload-ToFTP-WithRetry {
    param (
        [string]$LocalFile,
        [string]$FtpUrl,
        [string]$User,
        [string]$Pass,
        [int]$TimeoutSeconds = 15,
        [string]$RemoteFileName
    )
    
    # 检查本地文件是否存在
    if (-not (Test-Path $LocalFile)) {
        return @{
            Success = $false
            Message = "本地文件不存在"
            Details = ""
        }
    }
    
    # 构建FTP完整路径
    $RemotePath = "$FtpUrl/$RemoteFileName"
    
    # 读取本地文件的字节内容
    try {
        $FileBytes = [System.IO.File]::ReadAllBytes($LocalFile)
        $FileSize = $FileBytes.Length
    }
    catch {
        return @{
            Success = $false
            Message = "无法读取本地文件"
            Details = $_.Exception.Message
        }
    }
    
    # 尝试两种模式：先主动模式，后被动模式
    $ModesToTry = if ($PreferPassiveMode) {
        @($true, $false)  # 先被动，后主动
    } else {
        @($false, $true)  # 先主动，后被动
    }
    
    $LastError = ""
    
    foreach ($UsePassive in $ModesToTry) {
        $ModeName = if ($UsePassive) { "被动模式" } else { "主动模式" }
        
        try {
            # 创建FTP上传请求
            $FtpRequest = [System.Net.FtpWebRequest]::Create($RemotePath)
            $FtpRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
            $FtpRequest.Credentials = New-Object System.Net.NetworkCredential($User, $Pass)
            $FtpRequest.UseBinary = $true
            $FtpRequest.UsePassive = $UsePassive
            $FtpRequest.KeepAlive = $false
            $FtpRequest.Proxy = $null
            $FtpRequest.Timeout = $TimeoutSeconds * 1000
            $FtpRequest.ReadWriteTimeout = $TimeoutSeconds * 1000  # 添加读写超时
            
            # 写入数据到FTP
            $RequestStream = $FtpRequest.GetRequestStream()
            $RequestStream.Write($FileBytes, 0, $FileBytes.Length)
            $RequestStream.Close()
            $RequestStream.Dispose()
            
            # 获取响应
            $Response = $FtpRequest.GetResponse()
            $StatusCode = $Response.StatusCode
            $StatusDescription = $Response.StatusDescription
            $Response.Close()
            
            # 成功上传
            return @{
                Success = $true
                Message = "上传成功 [$ModeName]"
                Details = "状态: $StatusCode - $StatusDescription | 文件大小: $FileSize 字节"
            }
        }
        catch [System.Net.WebException] {
            $ErrorMsg = $_.Exception.Message
            $Status = $_.Exception.Status
            
            if ($Status -eq [System.Net.WebExceptionStatus]::Timeout) {
                $LastError = "[$ModeName] 超时 (${TimeoutSeconds}秒)"
            }
            elseif ($Status -eq [System.Net.WebExceptionStatus]::ConnectFailure) {
                $LastError = "[$ModeName] 连接失败: $ErrorMsg"
            }
            else {
                $LastError = "[$ModeName] 错误: $ErrorMsg"
            }
            
            # 如果是第一次尝试，继续下一个模式
            continue
        }
        catch {
            $LastError = "[$ModeName] 异常: $($_.Exception.Message)"
            continue
        }
    }
    
    # 两种模式都失败了
    return @{
        Success = $false
        Message = "上传失败（已尝试主动和被动模式）"
        Details = $LastError
    }
}

# -----------------------------------------------------------------------------
# 函数定义：启动异步上传任务
# -----------------------------------------------------------------------------
function Start-AsyncUpload {
    param (
        [string]$FilePath,
        [string]$FtpUrl,
        [string]$User,
        [string]$Pass,
        [int]$Timeout,
        [string]$RemoteFileName,
        [bool]$PreferPassive
    )
    
    # 如果有正在运行的上传任务，先检查状态
    if ($script:UploadJob -ne $null) {
        if ($script:UploadJob.State -eq 'Running') {
            return "SKIP"
        } else {
            Remove-Job -Job $script:UploadJob -Force -ErrorAction SilentlyContinue
            $script:UploadJob = $null
        }
    }
    
    # 创建上传任务的脚本块
    $UploadScriptBlock = {
        param($File, $Url, $Username, $Password, $Timeout, $OutputFileName, $PreferPassive)
        
        # 检查文件
        if (-not (Test-Path $File)) {
            return @{Success = $false; Message = "文件不存在"; Details = ""}
        }
        
        try {
            $FileBytes = [System.IO.File]::ReadAllBytes($File)
            $FileSize = $FileBytes.Length
        }
        catch {
            return @{Success = $false; Message = "读取文件失败"; Details = $_.Exception.Message}
        }
        
        $RemotePath = "$Url/$OutputFileName"
        
        # 尝试两种模式
        $ModesToTry = if ($PreferPassive) { @($true, $false) } else { @($false, $true) }
        $LastError = ""
        
        foreach ($UsePassive in $ModesToTry) {
            $ModeName = if ($UsePassive) { "被动" } else { "主动" }
            
            try {
                $FtpRequest = [System.Net.FtpWebRequest]::Create($RemotePath)
                $FtpRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
                $FtpRequest.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
                $FtpRequest.UseBinary = $true
                $FtpRequest.UsePassive = $UsePassive
                $FtpRequest.KeepAlive = $false
                $FtpRequest.Proxy = $null
                $FtpRequest.Timeout = $Timeout * 1000
                $FtpRequest.ReadWriteTimeout = $Timeout * 1000
                
                $RequestStream = $FtpRequest.GetRequestStream()
                $RequestStream.Write($FileBytes, 0, $FileBytes.Length)
                $RequestStream.Close()
                $RequestStream.Dispose()
                
                $Response = $FtpRequest.GetResponse()
                $StatusCode = $Response.StatusCode
                $Response.Close()
                
                return @{
                    Success = $true
                    Message = "上传成功 [${ModeName}模式]"
                    Details = "状态: $StatusCode | 大小: $FileSize 字节"
                }
            }
            catch {
                $LastError = "[${ModeName}模式] $($_.Exception.Message)"
                continue
            }
        }
        
        return @{Success = $false; Message = "两种模式均失败"; Details = $LastError}
    }
    
    # 启动后台任务
    $script:UploadJob = Start-Job -ScriptBlock $UploadScriptBlock -ArgumentList $FilePath, $FtpUrl, $User, $Pass, $Timeout, $RemoteFileName, $PreferPassive
    $script:LastUploadTime = Get-Date
    
    return "STARTED"
}

# -----------------------------------------------------------------------------
# 函数定义：检查上传任务状态
# -----------------------------------------------------------------------------
function Get-UploadStatus {
    if ($script:UploadJob -eq $null) {
        return $null
    }
    
    $JobState = $script:UploadJob.State
    
    if ($JobState -eq 'Completed') {
        $Result = Receive-Job -Job $script:UploadJob
        Remove-Job -Job $script:UploadJob -Force
        $script:UploadJob = $null
        return $Result
    }
    elseif ($JobState -eq 'Failed') {
        Remove-Job -Job $script:UploadJob -Force
        $script:UploadJob = $null
        return @{Success = $false; Message = "任务失败"; Details = ""}
    }
    elseif ($JobState -eq 'Running') {
        # 检查是否超时
        $ElapsedTime = (Get-Date) - $script:LastUploadTime
        if ($ElapsedTime.TotalSeconds -gt ($FtpTimeout * 2)) {
            Stop-Job -Job $script:UploadJob -ErrorAction SilentlyContinue
            Remove-Job -Job $script:UploadJob -Force -ErrorAction SilentlyContinue
            $script:UploadJob = $null
            return @{Success = $false; Message = "任务超时被强制终止"; Details = ""}
        }
        return "RUNNING"
    }
    
    return $null
}

# -----------------------------------------------------------------------------
# 主程序开始
# -----------------------------------------------------------------------------

Write-Host "`n" -NoNewline
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "     剪贴板监控脚本（本地优先 + 异步上传）v3.1" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# 步骤1：确保工作目录存在
Write-Host "`n[步骤1] 检查工作目录..." -ForegroundColor Yellow
if (-not (Test-Path $WorkDir)) {
    Write-Host "        工作目录不存在，正在创建: $WorkDir" -ForegroundColor Gray
    try {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        Write-Host "        ✓ 工作目录创建成功" -ForegroundColor Green
    }
    catch {
        Write-Host "        ✗ 工作目录创建失败: $_" -ForegroundColor Red
        Write-Host "`n程序退出" -ForegroundColor Red
        Read-Host "按回车键退出"
        exit
    }
} else {
    Write-Host "        ✓ 工作目录已存在: $WorkDir" -ForegroundColor Green
}

# 步骤2：构建完整的文件路径
$FullPath = Join-Path -Path $WorkDir -ChildPath $OutputFile
Write-Host "`n[步骤2] 日志文件路径: $FullPath" -ForegroundColor Yellow

# 步骤3：处理日志文件
Write-Host "`n[步骤3] 处理日志文件..." -ForegroundColor Yellow
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($ClearOldLog) {
    Write-Host "        清空模式已启用，正在创建新文件..." -ForegroundColor Gray
    try {
        [System.IO.File]::WriteAllText($FullPath, "", $Utf8NoBom)
        Write-Host "        ✓ 已清空旧日志，创建新文件" -ForegroundColor Green
    }
    catch {
        Write-Host "        ✗ 文件创建失败: $_" -ForegroundColor Red
        Read-Host "按回车键退出"
        exit
    }
} else {
    if (Test-Path $FullPath) {
        Write-Host "        ✓ 使用现有日志文件（追加模式）" -ForegroundColor Green
        $ExistingSize = (Get-Item $FullPath).Length
        Write-Host "        文件当前大小: $ExistingSize 字节" -ForegroundColor Gray
    } else {
        Write-Host "        日志文件不存在，正在创建..." -ForegroundColor Gray
        try {
            [System.IO.File]::WriteAllText($FullPath, "", $Utf8NoBom)
            Write-Host "        ✓ 日志文件创建成功" -ForegroundColor Green
        }
        catch {
            Write-Host "        ✗ 文件创建失败: $_" -ForegroundColor Red
            Read-Host "按回车键退出"
            exit
        }
    }
}

# 步骤4：初始化剪贴板状态
Write-Host "`n[步骤4] 初始化剪贴板..." -ForegroundColor Yellow
$PreviousClipboard = ""
try {
    $CurrentClipboard = Get-Clipboard -Raw -ErrorAction SilentlyContinue
    if ($CurrentClipboard) {
        $PreviousClipboard = $CurrentClipboard
        $PreviewText = if ($CurrentClipboard.Length -gt 30) { 
            $CurrentClipboard.Substring(0, 30) + "..." 
        } else { 
            $CurrentClipboard 
        }
        Write-Host "        ✓ 已读取当前剪贴板（此内容不会被保存）" -ForegroundColor Green
        Write-Host "        当前内容预览: $PreviewText" -ForegroundColor Gray
    } else {
        Write-Host "        ✓ 剪贴板为空" -ForegroundColor Green
    }
}
catch {
    Write-Host "        ⚠ 剪贴板读取失败，将从空白状态开始" -ForegroundColor Yellow
}

# 步骤5：测试FTP连接
Write-Host "`n[步骤5] 测试FTP连接..." -ForegroundColor Yellow
Write-Host "        正在测试连接到: $FtpServer" -ForegroundColor Gray
$TestResult = Upload-ToFTP-WithRetry -LocalFile $FullPath -FtpUrl $FtpServer -User $FtpUsername -Pass $FtpPassword -TimeoutSeconds $FtpTimeout -RemoteFileName $OutputFile

if ($TestResult.Success) {
    Write-Host "        ✓ FTP连接测试成功！" -ForegroundColor Green
    Write-Host "        $($TestResult.Details)" -ForegroundColor Gray
} else {
    Write-Host "        ⚠ FTP连接测试失败: $($TestResult.Message)" -ForegroundColor Yellow
    Write-Host "        $($TestResult.Details)" -ForegroundColor Gray
    Write-Host "        程序将继续运行，但上传功能可能不可用" -ForegroundColor Yellow
}

# 步骤6：显示运行信息
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "监控配置信息：" -ForegroundColor Cyan
Write-Host "  • 本地文件: $FullPath" -ForegroundColor White
Write-Host "  • FTP服务器: $FtpServer" -ForegroundColor White
Write-Host "  • 检查间隔: $Interval 秒" -ForegroundColor White
Write-Host "  • FTP超时: $FtpTimeout 秒" -ForegroundColor White
Write-Host "  • 上传模式: 异步（优先$(if ($PreferPassiveMode) {'被动'} else {'主动'})模式，自动切换）" -ForegroundColor White
Write-Host "  • 清空旧日志: $(if ($ClearOldLog) {'是'} else {'否'})" -ForegroundColor White
Write-Host "  • 编码格式: UTF-8 without BOM" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "`n⚠️  按 Ctrl+C 可停止监控`n" -ForegroundColor Yellow
Write-Host "开始监控剪贴板...\n" -ForegroundColor Green

# -----------------------------------------------------------------------------
# 主循环：监控剪贴板
# -----------------------------------------------------------------------------
$SaveCount = 0
$UploadCount = 0
$FailedCount = 0

try {
    while ($true) {
        # 检查上传任务状态
        if ($script:UploadJob -ne $null) {
            $UploadStatus = Get-UploadStatus
            if ($UploadStatus -ne $null -and $UploadStatus -ne "RUNNING") {
                if ($UploadStatus.Success) {
                    Write-Host "  [后台上传] ✓ $($UploadStatus.Message)" -ForegroundColor Green
                    if ($script:DebugMode -and $UploadStatus.Details) {
                        Write-Host "              $($UploadStatus.Details)" -ForegroundColor DarkGray
                    }
                    $UploadCount++
                } else {
                    Write-Host "  [后台上传] ✗ $($UploadStatus.Message)" -ForegroundColor Red
                    if ($UploadStatus.Details) {
                        Write-Host "              $($UploadStatus.Details)" -ForegroundColor DarkRed
                    }
                    $FailedCount++
                }
            }
        }
        
        # 监控剪贴板
        try {
            $CurrentClipboard = Get-Clipboard -Raw -ErrorAction SilentlyContinue
            
            if ($CurrentClipboard -and ($CurrentClipboard -ne $PreviousClipboard)) {
                
                $SaveCount++
                $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $ContentToSave = $CurrentClipboard + "`n"
                
                $Preview = if ($CurrentClipboard.Length -gt 60) { 
                    $CurrentClipboard.Substring(0, 60) + "..." 
                } else { 
                    $CurrentClipboard 
                }
                $Preview = $Preview -replace "`r`n", " " -replace "`n", " "
                
                Write-Host "────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
                Write-Host "[$Timestamp] 检测到新内容 (#$SaveCount)" -ForegroundColor Yellow
                Write-Host "内容预览: $Preview" -ForegroundColor White
                
                # 保存到本地
                Write-Host "→ 正在保存到本地..." -NoNewline -ForegroundColor Cyan
                try {
                    [System.IO.File]::AppendAllText($FullPath, $ContentToSave, $Utf8NoBom)
                    Write-Host " ✓ 本地保存成功" -ForegroundColor Green
                    
                    if (Test-Path $FullPath) {
                        $FileSize = (Get-Item $FullPath).Length
                        Write-Host "  本地文件大小: $FileSize 字节" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host " ✗ 本地保存失败: $_" -ForegroundColor Red
                    continue
                }
                
                # 异步上传到FTP
                Write-Host "→ 启动异步上传..." -NoNewline -ForegroundColor Cyan
                $StartResult = Start-AsyncUpload -FilePath $FullPath -FtpUrl $FtpServer -User $FtpUsername -Pass $FtpPassword -Timeout $FtpTimeout -RemoteFileName $OutputFile -PreferPassive $PreferPassiveMode
                
                if ($StartResult -eq "STARTED") {
                    Write-Host " ⏳ 后台上传中（不影响监控）" -ForegroundColor Yellow
                } elseif ($StartResult -eq "SKIP") {
                    Write-Host " ⏭ 跳过（上次上传还在进行）" -ForegroundColor Gray
                }
                
                $PreviousClipboard = $CurrentClipboard
                Write-Host ""
            }
        }
        catch {
            Write-Host "剪贴板读取异常: $_" -ForegroundColor Red
        }
        
        Start-Sleep -Seconds $Interval
    }
}
catch {
    Write-Host "`n`n" -NoNewline
}
finally {
    # 等待最后的上传任务
    if ($script:UploadJob -ne $null) {
        Write-Host "`n正在等待最后的上传任务完成..." -ForegroundColor Yellow
        $WaitCount = 0
        while ($script:UploadJob.State -eq 'Running' -and $WaitCount -lt 10) {
            Start-Sleep -Milliseconds 500
            $WaitCount++
        }
        
        if ($script:UploadJob.State -eq 'Completed') {
            $Result = Receive-Job -Job $script:UploadJob
            if ($Result.Success) {
                Write-Host "✓ 最后的上传任务已完成" -ForegroundColor Green
                $UploadCount++
            }
        }
        
        Remove-Job -Job $script:UploadJob -Force -ErrorAction SilentlyContinue
    }
    
    # 统计信息
    Write-Host "`n============================================================" -ForegroundColor Cyan
    Write-Host "监控已停止" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "统计信息：" -ForegroundColor White
    Write-Host "  • 本地保存: $SaveCount 条记录" -ForegroundColor White
    Write-Host "  • 上传成功: $UploadCount 次" -ForegroundColor Green
    Write-Host "  • 上传失败: $FailedCount 次" -ForegroundColor $(if ($FailedCount -gt 0) {'Red'} else {'Gray'})
    Write-Host "  • 本地文件: $FullPath" -ForegroundColor White
    
    if (Test-Path $FullPath) {
        $FinalSize = (Get-Item $FullPath).Length
        Write-Host "  • 文件大小: $FinalSize 字节" -ForegroundColor White
    }
    
    Write-Host "============================================================`n" -ForegroundColor Cyan
}
