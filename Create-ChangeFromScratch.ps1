#workflow Demo-CreateCRsForDemo
#{
    
    $creds = Get-AutomationPSCredential -Name 'SCSM Account'

    #Inlinescript{
        
        
            # Name:        Create-ChangeFromScratch
            # Description: Creates Change Requests in Service Manager
            # Author:      Michael Dugan
            # Date:        04/01/2016

            # Number of Change Requests to create
            $IncCount = 10
            # Delay between change request creation in seconds
            $Delay = 10

            # Import Modules
            Import-Module 'C:\Program Files\Common Files\SMLets\SMLets.psd1'
            Import-Module 'C:\ProgramData\SMARunbookContent\ChangeManagement.psm1'

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

            $Github = invoke-restmethod -uri https://api.github.com/repos/Azure/azure-powershell/issues

            $i = 1

            while ($i -le $IncCount)
            {
                $UserPool = @("")
                
                $CRUser = $UserPool | Get-Random
                
                # Define SCSM Relationship Users
                $CRAssignedUser = Get-SCSMObject -Class $CRUserClass | Where-Object {$_.DisplayName -eq $CRUser}
                $CRCreatedUser = Get-SCSMObject -Class $CRUserClass | Where-Object {$_.DisplayName -eq $CRUser}
                
                # Github API Objects
                $githubObjs = $Github
                
                # Github Title for CR
                $titleObj = $githubObjs | get-random | ? {! [string]::IsNullOrwhitespace($_.Title)}
                if ([String]$titleObj.Title -eq $null)
                {
                    throw "Title is null"
                }
                else
                {
                    $CRTitle = [String]$titleObj.Title
                }

                # Github Description for CR
                $descriptionObj = $githubObjs | get-random | ? {! [string]::IsNullOrwhitespace($_.Body)}
                if ([String]$descriptionObj.Body -eq $null)
                {
                    throw "Description is null"
                }
                else
                {
                    $CRDescription = [String]$descriptionObj.Body
                }

                # Github Reason for CR
                $reasonObj = $githubObjs | get-random | ? {! [string]::IsNullOrwhitespace($_.Title)}
                if ([String]$reasonObj.Title -eq $null)
                {
                    throw "Reason was not received from Github"
                }
                else
                {
                    $CRReason = [String]$reasonObj.Title
                }

                # Github Post Implementation Review for CR
                $postimplementationreviewObj = $githubObjs | get-random | ? {! [string]::IsNullOrwhitespace($_.Body)}
                if ([String]$postimplementationreviewObj.Body -eq $null)
                {
                    throw "Post Implementation Review was not received from Github"
                }
                else
                {
                    $CRPostImplementationReview = [String]$postimplementationreviewObj.Body
                }
                               
                $CRContactMethod = $CRUser | Get-Random
                
                $CurrentDate = Get-Date
                
                $StartDate = Get-Random -Minimum 0 -Maximum 365
                $CRStartDate = (Get-Date).AddDays($StartDate)
                                
                $EndDate = Get-Random -Minimum 1 -Maximum 7
                $CREndDate = $CRStartDate.AddDays($EndDate)
                
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
                $CREndDttm = $CRStartDttm.AddDays($CREndDate2)
                
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
                write-output "Calculating the Risk" 
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
                Write-Output "Determining the Change Type"
                $ChangeType = 'Minor'

                if ($RiskCount -ge 18)
                {
                    $ChangeType = 'Major Change Request'
                }

                if ($ChangeIsEmergency -eq 'True')
                {
                    $ChangeType = 'Emergency Change Request'
                }

                if ($ChangeIsImplemented -eq 'True')
                {
                    $ChangeType = 'Latent Change Request'
                }
                
                if ($RiskCount -ge 18)
                {
                    $ChangeType = 'Major Change Request'
                }

                    # Calculate date of the next CAB
                    $cstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time")
                    $dateutc = ($CurrentDate).ToUniversalTime()
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
                        $ChangeType = 'Emergency Change Request'
                        $ChangeIsEscalated = 'True'
                    }
                    elseif ($CREndDttm -gt $CABCutoffDttm -and $CREndDttm -lt $CABafterNextCABDttm)
                    {
                        Write-Output "Request Date '$($CRStartDttm)' is too close to the CAB ($NextCABDttm), Change is Emergency"
                        $ChangeType = 'Emergency Change Request'
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
                Write-Output "$RiskAssessmentPlanSummary"
                
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
                    $CRTitle = [String]$CRTitle.Substring(0,200)
                }
                else
                {
                    continue
                }
                if ($CRDescription.Length -gt 4000)
                {
                    $CRDescription = [String]$CRDescription.Substring(0,4000)
                }
                else
                {
                    continue
                }
                if ($CRReason.Length -gt 4000)
                {
                    $CRReason = [String]$CRReason.Substring(0,4000)
                }
                else
                {
                    continue
                }
                if ($CRPostImplementationReview.Length -gt 4000)
                {
                    $CRPostImplementationReview = [String]$CRPostImplementationReview.Substring(0,4000)
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
                                Risk = $Risk.DisplayName
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
                $MgmtPackTemplates = Get-SCSMObjectTemplate | ? {$_.ManagementPack -like 'CITSM.ChangeRequest.Library.CNCY'}
                
                # Retrieve Change Template
                switch -wildcard ($ChangeType)
                {
                    'Minor Change Request' {$template = $MgmtPackTemplates | ? {$_.DisplayName -eq 'Minor Change Request'}}
                    'Major Change Request' {$template = $MgmtPackTemplates | ? {$_.DisplayName -eq 'Major Change Request'}}
                    'Latent Change Request' {$template = $MgmtPackTemplates | ? {$_.DisplayName -eq 'Latent Change Request'}}
                    'Emergency Change Request' {$template = $MgmtPackTemplates | ? {$_.DisplayName -eq 'Emergency Change Request'}}
                    default {Write-Output "The template for '$ChangeType' Does Not Exist"}
                }
                
                $Template.ObjectCollection | ForEach-Object { Update-SCSMPropertyCollection $_ }
                
                Set-SCSMObjectTemplate -Projection $ChangeRequest -Template $Template
                $ChangeRequest = Get-SCSMObjectProjection System.WorkItem.ChangeRequestProjection -Filter "Id -eq '$Title'"
                
                # Create User Relationships
                Write-Output "Creating user relationships on CR"
                New-SCSMRelationshipObject -Relationship $AssignedRel -Source $CRObject -Target $CRAssignedUser -Bulk
                New-SCSMRelationshipObject -Relationship $CreatedRel -Source $CRObject -Target $CRCreatedUser -Bulk
                
                # Get Review Activities
                $WIContainsRAActivityRel = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity
                $AllCRActivities = Get-SCSMRelatedobject -SMObject $CRObject -Relationship $WIContainsActivityRARel
                $RActivities = $AllCRActivities | ? {$_.ClassName -eq "System.WorkItem.Activity.ReviewActivity"}
                Write-Output "Review Activity Count: $($RActivities.count)"

                # Get the Activity status 'In Progress' and set the first Activity to 'In Progress'
                $EnumStatus = Get-SCSMEnumeration -Name ActivityStatusEnum.Active 
                $FirstActivity = $AllCRActivities | where {$_.SequenceID -eq 0}
                Set-SCSMObject -SMObject $FirstActivity -Property Status -Value $EnumStatus 

                # Change ready for approval skips the first RA
                if (( 'Minor','Major' -contains $ChangeType) -and $ChangeReadyToApprove -eq $true)
                {
                    $RAToSkip = $AllCRActivities | ? {$_.SequenceId -eq 0}
                    Set-SCSMObject -SMObject $RAToSkip -Property 'skip' -Value 'True'
                }
                
                # Pause Before Creating Next Change Request
                Start-Sleep -Seconds $Delay
                $i += 1
                
            # Display End Time
            Write-Output "CR ($CRObject.Id)"    
            # Display End Time
            Write-Output "-------------------"
            $End = Get-Date
            Write-Output "Finished at $End"


    #End of Script
    #}-PSComputerName demo-scsm-01 -PSCredential $creds

#}
