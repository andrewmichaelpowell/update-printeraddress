# +++++++++++++++++++++++++++++++++++
# +  Update-PrinterAddress          +
# +  Author: Andrew Powell          +
# +  github.com/andrewmichaelpowell +
# +++++++++++++++++++++++++++++++++++

Function Update-PrinterAddress{
  Param(
    [Parameter(Mandatory="True",Position=0)]
    [string]$Computer,
    [Parameter(Mandatory="True",Position=1)]
    [String]$OldIP,
    [Parameter(Mandatory="True",Position=2)]
    [String]$NewIP
  )

  If(Test-Connection -ComputerName $Computer -Count 1 -Quiet){
    If($OldIP -Match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"){
      If($NewIP -Match "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"){
        Try{
          $Resolve = Get-PrinterPort -ErrorAction Stop -ComputerName $Computer -Name $OldIP
          Try{
            $Resolve = Get-PrinterPort -ErrorAction Stop -ComputerName $Computer -Name $NewIP
            Write-Host ""
            Write-Host -NoNewLine -ForegroundColor Yellow "Host "
            Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
            Write-Host -NoNewLine -ForegroundColor Yellow " already has a port "
            Write-Host -NoNewLine -ForegroundColor White $NewIP
            Write-Host -ForegroundColor Yellow "."
          }

          Catch{
            Add-PrinterPort -ComputerName $Computer -Name $NewIP -PrinterHostAddress $NewIP
            Get-Printer -ComputerName $Computer | Where-Object -Property PortName -eq $OldIP | Set-Printer -ComputerName $Computer -PortName $NewIP 
            Remove-PrinterPort -ComputerName $Computer -Name $OldIP

            Write-Host ""
            Write-Host -NoNewLine -ForegroundColor Yellow "Printer IP on host "
            Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
            Write-Host -NoNewLine -ForegroundColor Yellow " has been changed from "
            Write-Host -NoNewLine -ForegroundColor White $OldIP 
            Write-Host -NoNewLine -ForegroundColor Yellow " to "
            Write-Host -NoNewLine -ForegroundColor White $NewIP
            Write-Host -ForegroundColor Yellow "."
          }
        }

        Catch{
          Write-Host ""
          Write-Host -NoNewLine -ForegroundColor Yellow "Host "
          Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
          Write-Host -NoNewLine -ForegroundColor Yellow " does not have a port "
          Write-Host -NoNewLine -ForegroundColor White $OldIP
          Write-Host -ForegroundColor Yellow "."
        }
      }

      Else{
        Write-Host ""
        Write-Host -NoNewLine -ForegroundColor Yellow "Please enter a valid new IP address. "
        Write-Host -NoNewLine -ForegroundColor White $NewIP
        Write-Host -ForegroundColor Yellow " is not valid."
      }
    }

    Else{
      Write-Host ""
      Write-Host -NoNewLine -ForegroundColor Yellow "Please enter a valid old IP address. "
      Write-Host -NoNewLine -ForegroundColor White $OldIP
      Write-Host -ForegroundColor Yellow " is not valid."
    }
  }

  Else{
    Try{
      $Resolve = Resolve-DNSName -ErrorAction Stop -Name $Computer
      If(($Resolve.IPAddress).IndexOf(".") -gt 0){
        Write-Host ""
        Write-Host -NoNewLine -ForegroundColor Yellow "Host "
        Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
        Write-Host -ForegroundColor Yellow " has a DNS record, but it is currently offline."
      }

      Else{
        Write-Host ""
        Write-Host -NoNewLine -ForegroundColor Yellow "Host "
        Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
        Write-Host -ForegroundColor Yellow " does not have a DNS record."
      }
    }

    Catch{
      Write-Host ""
      Write-Host -NoNewLine -ForegroundColor Yellow "Host "
      Write-Host -NoNewLine -ForegroundColor White $Computer.ToLower()
      Write-Host -ForegroundColor Yellow " does not have a DNS record."
    }
  }
}
