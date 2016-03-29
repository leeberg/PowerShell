#workflow Demo-CreateCRsForDemo
#{
    
    $creds = Get-AutomationPSCredential -Name 'SCSM Account'

    #Inlinescript{
        
        
            # Name:        ChangeRequestCreationScript
            # Description: Creates Change Requests in Service Manager
            # Author:      Michael Dugan
            # Date:        03/29/2016

            # Number of Change Requests to create
            $IncCount = 10
            # Delay between change request creation in seconds
            $Delay = 10

            # Import SMLets
            Import-Module SMLets

            # Display Start Time
            $Start = get-date
            Write-Output "Started"
            Write-Output $Start
            Write-Output "-------------------`n"
            Write-Output "Created Change Request ID"
            Write-Output "-------------------"

            # Define SCSM User Class
            $CRUserClass = Get-SCSMClass -Name System.User$

            # Define Change Work Item Class
            $CRClass = Get-SCSMClass -Name System.WorkItem.ChangeRequest$

            # Define Change Enums
            $CRArea = Get-SCSMEnumeration -Name ChangeAreaEnum$
            $CRCategory = Get-SCSMEnumeration -Name ChangeCategoryEnum$
            $CRImpact = Get-SCSMEnumeration -Name ChangeImpactEnum$
            $CRImplementationResults = Get-SCSMEnumeration -Name ChangeImplementationResultsEnum$
            $CRPriority = Get-SCSMEnumeration -Name ChangePriorityEnum$
            $CRRisk = Get-SCSMEnumeration -Name ChangeRiskEnum$
            $CRStatus = Get-SCSMEnumeration -Name ChangeStatusEnum$

            # Define SCSM Relationship Variables
            $AssignedRel = Get-SCSMRelationshipClass -Name System.WorkItemAssignedToUser$
            $CreatedRel = Get-SCSMRelationshipClass -Name System.WorkItemCreatedByUser$
            
            # Define SCSM Relationship Users
            $CRAssignedUser = Get-SCSMObject -Class $CRUserClass | Where-Object {$_.DisplayName -eq $CRUser}
            $CRCreatedUser = Get-SCSMObject -Class $CRUserClass | Where-Object {$_.DisplayName -eq $CRUser}

            $i = 1

            while ($i -le $IncCount)
            {
                $UserPool = @("Duncan Lindquist", "John Hennen", "Bryan Schrippe", "Christopher Mank", "Guy Doggett", "John Hubert", "Lee Berg", "Matthew Selle", "Michael Dugan", "Nathan Lasnoski", "Rob Plank", "Ryan Ephgrave", "Steve Buchanan", "Steve Seibold", "Chiyo Odika", "Marcus Musial", "Matt Herman")
                
                $CRUser = $UserPool | Get-Random
                
                $CRTitle = (invoke-restmethod https://api.github.com/repos/Azure/azure-powershell/issues).title
                $CRTitle = [String]($CRTitle | ? {! [string]::IsNullOrwhitespace($_)} | Get-Random)

                $CRDescription = (invoke-restmethod https://api.github.com/repos/Azure/azure-powershell/issues).body
                $CRDescription = [String]($CRDescription | ? {! [string]::IsNullOrwhitespace($_)} | Get-Random)

                $CRReason = (invoke-restmethod https://api.github.com/repos/Microsoft/ChakraCore/issues).title
                $CRReason = [String]($CRReason | ? {! [string]::IsNullOrwhitespace($_)} | Get-Random)

                $CRPostImplementationReview = (invoke-restmethod https://api.github.com/repos/Microsoft/ChakraCore/issues).body
                $CRPostImplementationReview = [String]($CRPostImplementationReview | ? {! [string]::IsNullOrwhitespace($_)} | Get-Random)
                
                $CRContactMethod = $CRUser | Get-Random
                
                $CurrentDate = Get-Date
                
                $StartDate = Get-Random -Minimum 0 -Maximum 365
                $CRStartDate = (Get-Date).AddDays($StartDate)
                
                $EndDate = Get-Random -Minimum 1 -Maximum 7
                $CREndDate = (Get-Date).AddDays($StartDate+$EndDate)
                
                $Area = Get-SCSMChildEnumeration -Enumeration $CRArea | Get-Random
                
                $Category = Get-SCSMChildEnumeration -Enumeration $CRCategory | Get-Random
                
                $Impact = Get-SCSMChildEnumeration -Enumeration $CRImpact | Get-Random
                
                $ImplementationResults = Get-SCSMChildEnumeration -Enumeration $CRImplementationResults | Get-Random
                
                $Priority = Get-SCSMChildEnumeration -Enumeration $CRPriority | Get-Random
                
                $Risk = Get-SCSMChildEnumeration -Enumeration $CRRisk | Get-Random
                
                $Status = Get-SCSMChildEnumeration -Enumeration $CRStatus | Get-Random
                
                $CRImplementationPlan = "Standard Concurrency Process"
                
                $CRBackoutPlan = "Standard Concurrency Backout Process"
                
                $CRValidationPlan = "Stanard Concurrency Validation Process"
                
                $CRPostImpReviewContent = "The change was successfully implemented."
                                
                $ChangeIsEmergency = Get-SCSMChildEnumeration -Enumeration $CRCategory | Get-Random
                if ($ChangeIsEmergency.DisplayName -eq 'Emergency') {$ChangeIsEmergency = 'True'} else {$ChangeIsEmergency = 'False'}
                
                $ChangeIsImplemented = Get-SCSMChildEnumeration -Enumeration $CRImplementationResults | Get-Random
                if ($ChangeIsImplemented.DisplayName -eq 'Successfully Implemented') {$ChangeIsImplemented = 'True'} else {$ChangeIsImplemented = 'False'}
                
                $CRStartDate2 = Get-Random -Minimum 0 -Maximum 365
                $CRStartDttm = (Get-Date).AddDays($CRStartDate2)
                
                $CREndDate2 = Get-Random -Minimum 1 -Maximum 365
                $CREndDttm = (Get-Date).AddDays($CREndDate2)
                
                $RandomNumber = get-random -Minimum 1 -Maximum 8
                
                $RandomDate = (get-date).AddDays(-$RandomNumber)
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeRequiresDowntime = [bool]$Random
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeCausesDowntime = [bool]$Random
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeCanBackout = [bool]$Random
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeCanBeTested = [bool]$Random
                                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeValidationPlan = [bool]$Random
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeMakeStandard = [bool]$Random
                
                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeReadyToApprove = [bool]$Random

                $Random = Get-Random -Minimum 0 -Maximum 2
                $ChangeIsEscalated = [bool]$Random | Get-Random
                
                # Calculate Risk
                Write-Host "Calculating the risk" 
                $RiskCount = 0

                if ($ChangeRequiresDowntime -eq 'True')
                {
                    $RiskCount = $RiskCount + 18
                }
                if ($ChangeCausesDowntime -eq 'True')
                {
                    $RiskCount = $RiskCount + 18
                }
                if ($ChangeCanBackout -eq 'False')
                {
                    $RiskCount = $RiskCount + 9
                }
                if ($ChangeCanBeTested -eq 'False')
                {
                    $RiskCount = $RiskCount + 9
                }
                
                # Determine Change Type
                $ChangeType = 'Minor'

                if ($RiskCount -ge 18)
                {
                    $ChangeType = 'Major'
                }

                if ($ChangeIsEmergency -eq 'True')
                {
                    $ChangeType = 'Emergency'
                }

                if ($ChangeIsImplemented -eq 'True')
                {
                    $ChangeType = 'Latent'
                }

                # Retrieve Change Template
                $TemplateType = switch -wildcard ($ChangeType)
                {
                    'Minor' {$template = $MgmtPackTemplates | Where-Object {$_.DisplayName -eq 'Minor Change Request'}}
                    'Major' {$template = $MgmtPackTemplates | Where-Object {$_.DisplayName -eq 'Major Change Request'}}
                    'Latent' {$template = $MgmtPackTemplates | Where-Object {$_.DisplayName -eq 'Latent Change Request'}}
                    'Emergency' {$template = $MgmtPackTemplates | Where-Object {$_.DisplayName -eq 'Emergency Change Request'}}
                    default {Write-Output "The template for '$ChangeType' Does Not Exist"}
                }
                
                    #"Latent" {'Template.6268df6079e149ee8e680efc60fe2557'}
                    #"Emergency" {'ObjectTemplate.dcf987ea6f754ffd9d1fb924894f3f49'}
                    #"Minor" {'ObjectTemplate.8cb808af09394d68bb99f987ca5816c0'}
                    #"Major" {'ObjectTemplate.13381d026c5746029f55fc78df334e9d'}
                    #"Standard" {'Template.4fa9b664be554daaa37f55d9f4a00b0e'}
                    #Default {"Not a Valid Change Template Type!"}

                if ($RiskCount -ge 18)
                {
                    $ChangeType = 'Major'

                    # Calculate date of the next CAB
                    $cstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time")
                    $dateutc = ($CRStartDate).ToUniversalTime()
                    $csttime = [System.TimeZoneInfo]::ConvertTimeFromUtc($dateutc.ToUniversalTime(), $cstzone)
                    $datysOffset = [System.DayOfWeek]::Tuesday.value__ - $csttime.DayOfWeek.value__   #+ 7
                    if($datysOffset -lt 0 -or ($datysOffset -eq 0 -and $csttime.TimeOfDay.TotalHours -gt 8.0)){$datysOffset= $datysOffset +7}
                    $NextCABDttm = (Get-Date -Date $csttime -Hour 8 -Minute 0 -Second 0 -Millisecond 0).AddDays($datysOffset)
                    $CABafterNextCABDttm = $NextCABDttm.AddDays(7)
                    $CABCutoffDttm = $NextCABDttm.AddDays(-4)
                    Write-Output "Next CAB: $NextCABDttm "
                    Write-Output "CAB After Next CAB:$CABafterNextCABDttm"
                    Write-Output "CR Created Date '$($CRStartDate)'"
                    Write-Output "Scheduled Start Date '$($CRStartDttm)'"
                    
                    # Calculate EMERGENCY when requested too close to CAB meetings
                    if($CRStartDttm -lt $NextCABDttm)
                    {
                        Write-Output "Scheduled Start Date '$($CRStartDttm)' is before the CAB ($NextCABDttm), Change is Emergency"
                        $ChangeType = 'Emergency'
                        $ChangeIsEscalated = 'True'
                    }
                    elseif ($CREndDttm -gt $CABCutoffDttm -and $CREndDttm -lt $CABafterNextCABDttm)
                    {
                        Write-Output "Request Date '$($CRStartDttm)' is too close to the CAB ($NextCABDttm), Change is Emergency"
                        $ChangeType = 'Emergency'
                        $ChangeIsEscalated = 'True'
                    }
                    else
                    {
                        Write-Output 'Request dates do not indicate an Emergency'
                    }
                }

                Write-Output "Change type set to: $ChangeType"
                Write-Output "- Writing Risk Assessment Summary"
                $RiskAssessmentPlanSummary = ''
                $RiskAssessmentPlanSummary += "Change requires downtime to a critical service: $ChangeRequiresDowntime`n"
                $RiskAssessmentPlanSummary += "Failure of change would cause downtime to a critical service: $ChangeCausesDowntime`n"
                $RiskAssessmentPlanSummary += "Change can be backed out: $ChangeCanBackout`n"
                $RiskAssessmentPlanSummary += "Change can be tested: $ChangeCanBeTested`n"
                $RiskAssessmentPlanSummary += "Change is an emergency: $ChangeIsEmergency`n"
                $RiskAssessmentPlanSummary += "Major change is an Emergency due to requested dates: $ChangeIsEscalated`n"
                $RiskAssessmentPlanSummary += "Change is alredy implemented: $ChangeIsImplemented`n"
                $RiskAssessmentPlanSummary += "================================`n"
                $RiskAssessmentPlanSummary += "Calculated score: $ChangeType"
                
                # Post Implementation Review
                if ($ImplementationResults.DisplayName -eq "Successfully Implemented")
                {
                    $CRPostImplementationReview = $CRPostImpReviewContent
                }
                else
                {
                    continue
                }

                # Evaluate web request string lengths
                if ($CRTitle.length -gt 200)
                {
                    $CRTitle = $CRTitle.Substring(0,200)
                }
                else
                {
                    continue
                }
                if ($CRDescription.Length -gt 4000)
                {
                    $CRDescription = $CRDescription.Substring(0,4000)
                }
                else
                {
                    continue
                }
                if ($CRReason.Length -gt 4000)
                {
                    $CRReason = $CRReason.Substring(0,4000)
                }
                else
                {
                    continue
                }
                if ($CRPostImplementationReview.Length -gt 4000)
                {
                    $CRPostImplementationReview = $CRPostImplementationReview.Substring(0,4000)
                }
                else
                {
                    continue
                }
                
                $Params = @{
                                ID = "CR{0}"
                                Title = $CRTitle
                                Description = $CRDescription
                                Reason = $CRReason
                                Status = $Status.DisplayName
                                ContactMethod = $CRContactMethod
                                TestPlan = $CRValidationPlan
                                BackoutPlan = $CRBackoutPlan
                                ImplementationPlan = $CRImplementationPlan
                                ImplementationResults = $ImplementationResults.DisplayName
                                PostImplementationReview = $CRPostImplementationReview
                                Notes="ChangeComments:$ChangeComments`n`nStandard_Change_Selected:$ChangeMakeStandard`n`nChange_Ready_To_Approve:$ChangeReadyToApprove"
                                ScheduledStartDate = $CRStartDttm
                                ScheduledEndDate = $CREndDttm
                                RiskAssessmentPlan = "$RiskAssessmentPlanSummary"
                                CreatedDate = $CurrentDate.AddDays(-$RandomNumber)
                                Priority = $Priority.DisplayName
                                Impact = $Impact.DisplayName
                                Area = $Area.DisplayName
                                Risk = $CRRisk.DisplayName
                                Category = $Category.DisplayName
                                ScheduledDowntimeStartDate = $RandomDate
                                ScheduledDowntimeEndDate = $RandomDate
                                ActualDowntimeStartDate = $RandomDate
                                ActualDowntimeEndDate = $RandomDate
                                }
                
                # Create The New Object
                Write-Output "Creating a New CR"
                $CRObject = New-SCSMObject -Class $CRClass -PropertyHashTable $Params -pass
                
                Write-Output "CR ID is: $($CRObject.Id)"
                
                $Title = $CRObject.Id
                $ChangeRequest = Get-SCSMObjectProjection System.WorkItem.ChangeRequestProjection -Filter "Id -eq '$Title'"

                Write-Output "Getting Management Pack Templates"
                $MgmtPackTemplates = Get-SCSMObjectTemplate | Where-Object {$_.ManagementPack -like 'CITSM.ChangeRequest.Library'}
                
                Write-Output "Setting Management Pack Template on CR"
                $Template = Get-SCSMObjectTemplate $TemplateType
                
                $Template.ObjectCollection | ForEach-Object { Update-SCSMPropertyCollection $_ }
                
                Set-SCSMObjectTemplate -Projection $ChangeRequest -Template $Template
                $ChangeRequest = Get-SCSMObjectProjection System.WorkItem.ChangeRequestProjection -Filter "Id -eq '$Title'"
                
                # Create User Relationships
                Write-Output "Creating user relationships on CR"
                New-SCSMRelationshipObject -Relationship $AssignedRel -Source $CRObject -Target $CRAssignedUser -Bulk
                New-SCSMRelationshipObject -Relationship $CreatedRel -Source $CRObject -Target $CRCreatedUser -Bulk
                
                # Pause Before Creating Next Change Request
                Start-Sleep -Seconds $Delay
                $i += 1
            }
                
            # Display End Time
            Write-Output "-------------------`n"
            $End = get-date
            Write-Output "Finished"
            Write-Output $End

    #End of Script
    #}-PSComputerName demo-scsm-01 -PSCredential $creds

#}