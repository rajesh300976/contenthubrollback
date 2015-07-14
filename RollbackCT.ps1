# Created by GANESAN RAJESH KUMAR

# Load IIS module:
Import-Module WebAdministration

$XmlFilePath="Solutions.xml"


function GetHelp() {
$Helptext = @"

DESCRIPTION:
NAME: RollbackCT.ps1
Removes the ContentType listed in Solutions.xml to the site input as a parameter


Displays the help topic for the script

"@

$helpText;
}

Function AddSnapIn()
{
    #handles exceptions caused by trying to add a snapin
	Trap [Exception]
	{
	     
	      continue; 
	}

    #Check that the required snapins are available , use a comma delimited list.
    #example
    # ("Microsoft.SharePoint.PowerShell", "Microsoft.Office.Excel")
	$RequiredSnapIns = ("Microsoft.SharePoint.PowerShell");
	ForEach ($SnapIn in $RequiredSnapIns)
	{
		if ( (Get-PSSnapin -Name $SnapIn -ErrorAction SilentlyContinue) -eq $null ) 
		{ 
		    Add-PsSnapin $SnapIn
		} 
	}
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint");
}
#—————————————————————————-
# Delete Field
#—————————————————————————-
function DeleteField([string]$siteUrl, [string]$fieldName) {
    #Write-Host “Start removing field:” $fieldName -ForegroundColor "green"
    $site = Get-SPSite $siteUrl
    $web = $site.RootWeb

	if($web.Fields.ContainsFieldWithStaticName($fieldName))
	{
    #Delete field from all content types
    foreach($ct in $web.ContentTypes) {
        $fieldInUse = $ct.FieldLinks | Where {$_.Name -eq $fieldName }
        if($fieldInUse) {
            Write-Host "Remove field from CType:" $ct.Name -ForegroundColor "green"
            $ct.FieldLinks.Delete($fieldName)
            $ct.Update()
			Write-Host "Completed removing field from CType :" $fieldName -ForegroundColor "green"
        }
    }

    #Delete column from all lists in all sites of a site collection
    $site | Get-SPWeb -Limit all | ForEach-Object {
       #Specify list which contains the column
        $numberOfLists = $_.Lists.Count
        for($i=0; $i -lt $_.Lists.Count ; $i++) {
            $list = $_.Lists[$i]
            #Specify column to be deleted
			if($list)
			{
            #if($list.Fields.ContainsFieldWithStaticName($fieldName)) {
			try
			{
                $fieldInList = $list.Fields.GetFieldByInternalName($fieldName)
			
                if($fieldInList) {
                    Write-Host “Delete column from ” $list.Title ” list on:” $_.URL -ForegroundColor "green"

                 #Allow column to be deleted
                 $fieldInList.AllowDeletion = $true
                 #Delete the column
                 $fieldInList.Delete()
                 #Update the list
                 $list.Update()
                }
				}
			catch
			{
				#Write-Host "."
			}
            #}
        }}
    }

    # Remove the field itself
    if($web.Fields.GetFieldbyinternalname($fieldName)) {
        Write-Host “Remove field:” $fieldName -ForegroundColor "green"
        $web.Fields.Delete($fieldName)
    }
}
    $web.Dispose()
    $site.Dispose()
}



function DeleteColumn
{
Param ([string]$FieldName)

#Declare the name of the Field/Column to be deleted
$FieldTobeDeleted = $FieldName
#Iterating webs....

$siteCollCnt = 0
$siteCnt = 0
$webAppCnt = 0
$siteCntAll=0
$siteCollCntAll=0
try
{
$AllWebApps = Get-SPWebApplication
foreach($webApp in $AllWebApps)
{

$siteCollCnt = 0
$webAppCnt++ 
	if($webApp -ne $null)
	{
	  Write-Host "Web Application : "  $webApp.Name

	   foreach($siteColl in $webApp.Sites)
		{ 
		$siteCollCnt++
		$siteCollCntAll++
		if($siteColl -ne $null ) 
		  {

			 Write-Host "Checking Site Collection : " $siteColl.Url -foregroundcolor "green"	
			try
			{
			$site = Get-SPWeb $siteColl.Url
			}
			catch
			{
			$site = $null
			}
			if($site -eq $null)
			{
			Write-Host "[INFO]: Site not accessible.. Skipping Sites"
			}
			else
			{
			try{
			$siteCnt = 0
				#foreach($web in $siteColl.AllWebs)
				foreach ($web in $site.Site.AllWebs)
				{       
				$siteCnt++
				$siteCntAll++
				Write-Host "Site : " $web.Url
					#Stage 1: Removing Field from Site collection's Content types
					if ($web.IsRootWeb)
					{																								
						
						foreach($ct in $web.ContentTypes) 
						{
						   try
						   {
							$field = $ct.FieldLinks[$FieldTobeDeleted]			
							if($field -ne $null) {	
								#Write-Host "Field Found in Content Type: Printing the field ID and deleting the field from content type" $ct.Name "and Field ID"  $ct.FieldLinks[$FieldTobeDeleted].Id -foregroundcolor "yellow"	

								$ct.set_readOnly($false); 
								$ct.FieldLinks.Delete($FieldTobeDeleted)
								$ct.Update()
							}
							}
							catch
							{
							Write-Host "[Error]: Deleting field from content type: " $_.exception.message $_.exception.ItemName 
							}
						}
						#Stage 4: Removing Field from Site columns
						try
						{
						$field = $web.Fields.getFieldbyInternalName($FieldTobeDeleted)
						}
						catch
						{
						$field = $null
						}
						try
						{
							if($field) 
							{
								#Write-Host "Field Found in Site Column: Printing the field ID and deleting the field"  $field.Id  $field.Name -foregroundcolor "yellow"	
			
								$web.Fields.Delete($FieldTobeDeleted)
								Write-Host "[Deletecoulmn] : $FieldTobeDeleted " $web.Url
							}
						}
						catch
						{
						 Write-Host "[Error]:Deleting site column" $_.exception.message ":" $_.exception.ItemName
						}
					}	

					$web.Dispose();
				}
				Write-Host "Sites Count for " $siteColl.Url " : " $siteCnt -foregroundcolor "green"
				$siteColl.CatchAccessDeniedException = $true; 
				$siteColl.Dispose(); 
				}
				catch
				{
				Write-Host "Error : " $_.Exception.Message " : Error Item : " $_.Exception.ItemName
				
				}}
		  #} end if
		  
		  }
		}
		Write-Host "Site Collection Count for " $webApp.Name " : " $siteCollCnt -foregroundcolor "green"
	   } 
 }  
 Write-Host "*********************************"
 Write-Host "Total WebApplications Count : " $webAppCnt -foregroundcolor "green"
 Write-Host "Total Sitecollection count : " $siteCollCntAll -foregroundcolor "green"
 Write-Host "Total Sites count : " $siteCntAll -foregroundcolor "green"
 Write-Host "*********************************"
 
 
 }
   catch{
   Write-Host "[Error]: Get Web Aplication" $_.exception.message " : " $_.exception.itemName 
 }
 
 
}

function RemoveContentTypeFromSubscribtionSite
{
Param ([string]$Url)
$webApp = Get-SPWebApplication $Url
#Iterating webs....
if($webApp -ne $null)
{
  Write-Host "Web Application : "  $webApp.Name

   foreach($siteColl in $webApp.Sites)
    { 
	
    if($siteColl -ne $null ) 
      {
	  if($siteColl.Url -like "*/cthubs/*")
	  {
	  }
	  else
	  {
	    
         Write-Host "Site Collection : " $siteColl.Url -foregroundcolor "green"	


            RemoveContentType $siteColl.Url
            $siteColl.CatchAccessDeniedException = $true; 
            $siteColl.Dispose(); 
      }
	  }
    }
   } 
}

function Get-ScriptLogFileName([string]$loggingAction)
{
              $currentDate = Get-Date -Format "ddMMMyyyyHHmm"
              $logFileName = $loggingAction + $currentDate  + ".log"

              Return $logFileName
} 

Function RemoveContentType
{
Param ([string]$Url)

$site  = get-spsite $Url
$web = $site.RootWeb

$ctypeName = "Your Content Type Name"

$ct = $web.contenttypes[$ctypeName]
if($ct)
{
$ct.readonly=$false
$ct.update();



$ctusage = [Microsoft.SharePoint.SPContentTypeUsage]::GetUsages($ct)
foreach($ctuse in $ctusage)
{

$list = $web.GetList($ctuse.Url)
if($list)
{
$contenttypeCollection = $list.ContentTypes
$contenttypeCollection[$ctypeName].readonly=$false
$contenttypeCollection.delete($contenttypeCollection[$ctypeName].Id);
$contenttypeCollection[$ctypeName].update();
Write-host "Deleted $ctypeName from $Url"
}
}
$ct.delete();
$web.update()
Write-host "Deleted $ctypeName successfully from $Url" 
}
else
{
Write-host "Cannot find $ctypeName.. Skipping delete"
}

$web.dispose()

}

function StartJobOnWebApp
{
    param([string]$WebAppName, [string]$JobName)
  
    $WebApp = Get-SPWebApplication $WebAppName;
  

    ##Getting right job for right web application
    $job = Get-SPTimerJob | ?{$_.Name -match $JobName} | ?{$_.Parent -eq $WebApp}
    if($null -ne $job)
    {
        $startet = $job.LastRunTime
		Write-Host "[Timerjob]: ContentType Subscriber Started..." -foregroundcolor "green"	
        #Write-Host -ForegroundColor Yellow -NoNewLine "Running"$job.DisplayName"Timer Job."
        Start-SPTimerJob $job

        ##Waiting til job is finished
        while (($startet) -eq $job.LastRunTime)
        {
            Write-Host -NoNewLine -ForegroundColor Yellow "."
            Start-Sleep -Seconds 2
        }

        ##Checking for error messages, assuming there will be errormessage if job fails
        if($job.ErrorMessage)
        {
            Write-Host -ForegroundColor Red "Possible error in" $job.DisplayName "timer job:";
            Write-Host "LastRunTime:" $lastRun.Status;
            Write-Host "Errormessage:" $lastRun.EndTime;

        }
        else
        {
            Write-Host -ForegroundColor Green $job.DisplayName"Timer Job has completed.";
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "ERROR: Timer job" $job.DisplayName "on web application" $WebApp "not found."
    }

}

Function UnPublish-ContentTypeHub {     
param    (         [parameter(mandatory=$true)][string]$CTHUrl,         [parameter(mandatory=$true)][string]$Group    )       
$site = Get-SPSite $CTHUrl    
if(!($site -eq $null))     
{         
$contentTypePublisher = New-Object Microsoft.SharePoint.Taxonomy.ContentTypeSync.ContentTypePublisher ($site)         
$site.RootWeb.ContentTypes | ? {$_.Group -match $Group} | % {  
if($_.Name -match "Your Content Type")
{$contentTypePublisher.UnPublish($_)             
write-host "Content type" $_.Name "has been Unpublished" -foregroundcolor Green     
    } 
}    
} 
}

if($help) { GetHelp; Continue }
if($XmlFilePath) 
{ 
    AddSnapIn;
      
	try
	{
	stop-transcript|out-null
	}
	catch
	[System.InvalidOperationException]{}	
	try
	{
	$Logfile = Get-ScriptLogFileName("RollbackCT_Results")  
    Start-Transcript -Path $Logfile -Force
	$currentDateTime = Get-Date -Format "dd-MMM-yyyy HH:mm"
	
    
    [xml]$s = get-content $XmlFilePath   

    #for each solution item in the xml file...
	$SiteCollectionUrl = $s.Configuration.CTHub.url;
    Write-Host $SiteCollectionUrl;
	Write-Host "ContentType Rollback Started"
	
	Write-Host "ContentType UnPublish Started" -foregroundcolor "green"
	UnPublish-ContentTypeHub $SiteCollectionUrl "Content Type Group"
	
	#Delete the SiteColumns from the content hub
	DeleteField $SiteCollectionUrl "Enter the internal field name"
	DeleteField $SiteCollectionUrl "Enter the internal field name"
	DeleteField $SiteCollectionUrl "Enter the internal field name"
	DeleteField $SiteCollectionUrl "Enter the internal field name"
	DeleteField $SiteCollectionUrl "Enter the internal field name"
	

	RemoveContentType $SiteCollectionUrl

   

	$WebAppUrl = $s.Configuration.WebApplication.url;

	Write-Host "[Timerjob]: ContentType Hub Started..." -foregroundcolor "green"	
	#Run the Content Type Hub timer job
	$ctHubTJ = Get-SPTimerJob "MetadataHubTimerJob" 
	$ctHubTJ.RunNow() 

	Start-Sleep -s 30
	
	#Run the Content Type Subscriber timer job for a specific Web Application
	StartJobOnWebApp $WebAppUrl "MetadataSubscriberTimerJob"
	
	Write-Host "ContentType Publish Completed" -foregroundcolor "green"
	
    Write-Host "Deletion of Site Column from Subscribed Site Collections Started" -foregroundcolor "green"	
    RemoveContentTypeFromSubscribtionSite $WebAppUrl
	DeleteColumn  <Enter Internal Field Name>
	DeleteColumn  <Enter Internal Field Name>
	DeleteColumn  <Enter Internal Field Name>
	DeleteColumn  <Enter Internal Field Name>
	DeleteColumn  <Enter Internal Field Name>
	
	
	Write-Host "Deletion of Site Column from Site Collection End" -foregroundcolor "green"	

	#Write-Host " Done."
	Write-Host "ContentType Rollback completed"  -foregroundcolor "green"
    
		}
	catch
	{
	Write-Host "Error : RollbackCTHub : " $_.exception.message " : " $_.exception.ItemName
	}
	finally
	{
	Stop-Transcript
	
	$log = Get-Content $Logfile
    $log > $Logfile
	}
}
else
{
    GetHelp;
}
