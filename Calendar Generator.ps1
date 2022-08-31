$serving_days = Get-Date -Year 2022 -Month 3 -Day 21 


$end = Get-Date -Year 2023 -Month 3 -Day 20


$calendar = @()

do{
    
    $serving_days = $serving_days.AddDays(1) 

    
    
    $calendar += Get-Date $serving_days -Format "dddd MM/dd/yyyy"
    

}


while($serving_days -le $end)


$leave_object = @()

foreach($day in $calendar)
{

    
        $leave_object += New-Object psobject -Property @{

        "DOW" = $day.Split()[0]

        "Date" = $day.Split()[1]}
    
    
    
}
    
   


$leave_object |Export-Excel $scriptPath\Leaves.xlsx