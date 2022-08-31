#Get the script location   
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

function reset-leave
{
    $obj = Import-Excel $scriptPath\Leaves.xlsx

    $updated = @()

    foreach($day in $obj)
    {
        $updated += New-Object psobject -Property @{
                   
                    "Date" = $day.Date
                    "DOW" = $day.DOW
                    "Apply" = 0 }
    }

    $updated |Export-Excel $scriptPath\Leaves.xlsx

    update-leave
}

function apply-leave
{
    param($date)

    $obj = Import-Excel $scriptPath\Leaves.xlsx

    $updated =@()
    
    foreach($day in $obj)
    {
        
        if($day.Date -eq $date){
            
            $updated += New-Object psobject -Property @{
                   
                    "Date" = $day.Date
                    "DOW" = $day.DOW
                    "Apply" = 1
            
            }
        }
        
        else{

            $updated += New-Object psobject -Property @{
                   
                    "Date" = $day.Date
                    "DOW" = $day.DOW
                    "Apply" = $day.Apply
        
            }
        }
        
        
        
        
    }

        $updated |Export-Excel $scriptPath\Leaves.xlsx

        update-leave
}

function update-leave
{
    
    $updated = @()
    
    $obj = Import-Excel $scriptPath\Leaves.xlsx

    $balance = 0
    $total_AL = 12

    foreach($day in $obj){

        $balance = $balance + ($total_AL/365) - $day.Apply

        $updated += New-Object psobject -Property @{
                   
                    "Date" = $day.Date
                    "DOW" = $day.DOW
                    "Apply" = $day.Apply
                    "Balance" =[math]::Round($balance,1)
        
            }

    }


        $updated |Export-Excel $scriptPath\Leaves.xlsx
    

}


do{
    cls

    Write-Host "Leave Buddy v1.0 :)"
    Write-Host
    Write-Host "1.Apply Leave"
    Write-Host "2.Show Leaves Status"
    Write-Host "3.Show Working Calendar"
    Write-Host "4.Cancel Leave"
    Write-Host "5.Quit"
    write-host
    $choice = Read-Host "Enter option and press Enter"

    if($choice -eq 1){
    
        cls
        
        $date = Read-Host "Enter date"

        apply-leave $date
    }

    elseif($choice -eq 2){

        cls
        
        $excel = Import-Excel $scriptPath\Leaves.xlsx
        
        $obj = @()
        
        
        foreach($day in $excel){
            
            $obj += New-Object psobject -Property @{
            
                    "Date" = $day.Date
                    "DOW" = $day.DOW
                    "Apply" = $day.Apply
                    "Balance" = $day.Balance
        
        }
        
        
        }
        
        
        
        
        $obj|Where-Object{$_.apply -eq 1}|Out-Default        
        
        Write-Host "Total Annual Leaves: 12 days"
        
        Write-Host "Total applied leaves:" ($obj|Where-Object{$_.apply -eq 1}).count "days"

        Write-Host "Balance as of today:" (($obj|Where-Object{$_.date -eq (get-date -Format MM/dd/yyyy)}).balance) "days"

        Write-Host "Balance as of last day:" ($obj[364]).Balance "days"

        Write-Host
        
        read-host "Press Enter to go back"
    }

    
    elseif($choice -eq 3){
    }
    
    
    
    
    
    elseif($choice -eq 5){
        
        cls
        
        break
    }
}

while($true)