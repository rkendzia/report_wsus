<#

Script écrit par raphaël Kendzia

Ce script fait l'inventaire d'une infrastructure RDS. Il permet également de vérifier les principaux 
points pour faire du troubleshooting.



#>

param(
    [string]$Client
    )



#region Découverte des collections, Serveurs ...
$ErrorActionPreference = "SilentlyContinue"
$NameServer= [System.Environment]::MachineName
$Status = (Get-WsusServer).GetStatus()
$ErrorComputer = @(Get-WsusComputer -IncludedInstallationStates Failed)
$SuccessComputer =  Get-WsusComputer -IncludedInstallationStates Installed
$PendingReboot =  Get-WsusComputer -IncludedInstallationStates InstalledPendingReboot
$NotInstalled =  Get-WsusComputer -IncludedInstallationStates NotInstalled
$test = Get-WsusComputer -IncludedInstallationStates 
$Dowloaded =  Get-WsusComputer -IncludedInstallationStates Downloaded
$ListCategory = (Get-WsusServer).GetSubscription().GetUpdateCategories() 
$ListClassification = (Get-WsusServer).GetSubscription().GetUpdateClassifications() 
$ListGroups = (Get-WsusServer).GetComputerTargetGroups() 
$ErrorUpdates = @(Get-WsusUpdate -Approval Approved -Status Failed | select  @{Label="Update";Expression={$_.Update.Title}} )




#endregion











if ($htmlfile -eq $null ) {$htmlfile= ".\Rapport Wsus - "+ $Client+".html" }

#region head HTML
$HeadHTML=('
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewporst" content="width=device-width, initial-scale=1.0">
    <title>Rapport DC </title>
    <style>
        h1 {

           color: rgba(15, 8, 119, 0.76);
        }
        h2 {

            color: blue;
         }
         h3 {

            color: #1E90FF;
         }
         h4 {

            color: #7B68EE;
         }
        TABLE {TABLE-LAYOUT: fixed; border:0.1 solid gray ; FONT-WEIGHT: normal; FONT-SIZE: 8pt;  FONT-FAMILY: Tahoma; ; border : 1}
		td{ VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma ;TEXT-ALIGN: center;border:0.1 solid gray ;padding 0px 0px}
		th {VERTICAL-ALIGN: TOP; TEXT-ALIGN: center;BACKGROUND-COLOR: #0066FF;COLOR: white ;}
				
		tr{border:0.1; padding 0px 0px}
		tr:nth-child(even) {background: #6dc2e9}
		tr:nth-child(odd) {background: #FFF}
          }  
        body {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 8pt; RIGHT: 10px; COLOR: Black; FONT-FAMILY: Tahoma; POSITION: absolute;}
        p.ok {FONT-FAMILY: Tahoma; FONT-SIZE: 10pt;COLOR:green}
        p.err {FONT-FAMILY: Tahoma; FONT-SIZE: 10pt;COLOR:red}
        
            


    </style>
</head>
<div style="text-align:center;background-color: #0B67AB">
 <div style="display:inline-block; text-align:left; font-size:12pt;
             background-color:white; 
             padding: 20px;
             border:2px solid rgb(0,0,0);
             width: 60%;
             min-height: 1000px;">
 

<body>
    <h1>Rapport Wsus de '+$Client+'</h1>


    
    Date du rapport '+(get-date -Format D) +' <br>
')


#endregion









#region Informations   
$BodyHtml += @("

    <H2>Informations du serveur</H2>
    <table >
        <tr>
            <th>Nom du serveur</th>
            <td>"+$NameServer+"</td>
        </tr>
        <tr>
            <th>Port utilisé</th>
            <td>"+$(Get-WsusServer |select -ExpandProperty PortNumber)+"</td>
        </tr>
        <tr>
            <th>Version</th>
            <td>"+$(Get-WsusServer | select -ExpandProperty Version)+"</td>
        </tr>
        <tr>
            <th>URL</th>
            <td>"+$(Get-WsusServer | select -ExpandProperty WebServiceUrl)+"</td>
        </tr>
        
    </table>


    ")    
   
#endregion



#region Status du Serveur   
$BodyHtml += @("

    <H2>Status</H2>
    <table >
        <tr>
            <th>Nombre de mise à  jour</th>
            <td>"+$Status.updatecount+"</td>
        </tr>
        <tr>
            <th>Mise à  jour déclinée</th>
            <td>"+$Status.Declinedupdatecount+"</td>
        </tr>
        <tr>
            <th>Mise à  jour approuvée</th>
            <td>"+$Status.ApprovedUpdateCount+"</td>
        </tr>
        <tr>
            <th>Mise à  jour non approuvée</th>
            <td>"+$Status.NotApprovedUpdateCount+"</td>
        </tr>
        <tr>
            <th>Nombre de mise à  jour en erreur</th>
            <td p.err>"+$Status.UpdatesWithClientErrorsCount+"</td>
        </tr>
        <tr>
            <th>Nombre de mise à  jour nécessaire</th>
            <td>"+$Status.UpdatesNeededByComputersCount+"</td>
        </tr>
        <tr>
            <th>Nombre de de clients</th>
            <td>"+$Status.ComputerTargetCount+"</td>
        </tr>
         <tr>
            <th>Nombre de de clients avec mise à  jour en attente</th>
            <td>"+$Status.ComputerTargetsNeedingUpdatesCount+"</td>
        </tr>
        <tr>
            <th>Nombre de de clients avec des mises à  jour en erreur</th>
            <td p.err>"+$Status.ComputerTargetsWithUpdateErrorsCount+"</td>
        </tr>
         <tr>
            <th>Nombre de de clients à  jour</th>
            <td>"+$Status.ComputersUpToDateCount+"</td>
        </tr>
        
    </table>


    ")    
   
#endregion


#region espace disque

$BodyHtml += @("

<H2>Vérification de l'espace dique</H2>
")


              $temp += @("
           
            <table border='1'>
                <tr>
                    <th>Nom du disque</th>
                    <th>Taille</th>
                    <th>Occupation</th>
                    <th>libre</th>
                </tr>
                ")
     $BodyHtml += $temp
     $temp =$null
       $Disque = Get-WmiObject -Class Win32_LogicalDisk  | where {$_.drivetype -eq "3"}
        

            $Disque | % { 
		    $Nom = $_.DeviceID
            $type = $_.DriveType
		    $Taille = [math]::round($_.Size/ 1gb,2)
		    $EspaceLibre = [math]::round($_.Freespace / 1gb,2)
		    $Occupation = [math]::round(($Taille - $EspaceLibre) *100 / $Taille,2)
		    $Libre = [math]::round(($EspaceLibre / $Taille)*100,2)
            
            
                
		        
                
		            if ($Libre -lt 10) {
                        
                        $temp +="<tr><td>"+$Nom+"</td><td>"+$Taille+"</td><td>"+$Occupation+"</td><td bgcolor='red'>"+$Libre+"</td></tr>"
                        $BodyHtml += $temp
                        $temp =$null
			            
		             }else{
			                $temp +="<tr><td>"+$Nom+"</td><td>"+$Taille+"</td><td>"+$Occupation+"</td><td bgcolor='green'>"+$Libre+"</td></tr>"
                            $BodyHtml += $temp
                            $temp =$null

                        }
           
    }
    
        $BodyHtml += "</table>"
    

    
#endregion





#region Serveur Parent
$checkSerPar= (Get-WsusServer).GetConfiguration() | select -ExpandProperty UpstreamWsusServerName
if ( $checkSerPar )
{
    $ConfigServerParent = $((Get-WsusServer).GetConfiguration())
    $NameServerParent = $ConfigServerParent.UpstreamWsusServerName
    $PortServerParent = $ConfigServerParent.UpstreamWsusServerPortNumber
    $SSLServerParent = $ConfigServerParent.UpstreamWsusServerUseSsl
    $BodyHtml += @("

        <H2>Informations du serveur Parent</H2>
        <table >
            <tr>
                <th>Nom du serveur</th>
                <td>"+$NameServerParent+"</td>
            </tr>
            <tr>
                <th>Port utilisé</th>
                <td>"+$PortServerParent+"</td>
            </tr>
            <tr>
                <th>SSL</th>
                <td>"+$SSLServerParent+"</td>
            </tr>
           
        
        </table>

        ")
}



    


#endregion




#region Classifications

$BodyHtml += @("
    
    <H2>Classification</H2>
    <p>Voici la liste des classifications configurées sur le serveur.

    <table>
        <tr>
            <th>Liste</th>
         </tr>
    ") 
    foreach ( $listClass in $ListClassification){
        $NameClassification = $listClass.title
        $temp +=   "<tr><td>"+$NameClassification+"</td></tr>"


        }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
#endregion


#region Category

$BodyHtml += @("
    
    <H2>Catégories</H2>

    <table>
        <tr>
            <th>Liste</th>
         </tr>
    ") 
    foreach ( $listCat in $ListCategory){
        $NameCat = $listCat.title
        $temp +=   "<tr><td>"+$NameCat+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
#endregion


#region groupes

$BodyHtml += @("
    
    <H2>Groupes</H2>

    <table>
         <tr>
            <th>Liste</th>
            <th>groupe parent</th>
            <th><Nombre de poste</th>
         </tr>
    ") 
    foreach ( $listgr in $ListGroups)
    {
        $NameGroupes = $listgr.name
        $IDGroupes = $listgr.Id
        $IDGroupes
        
        $temp +=   "<tr><td>"+$NameGroupes+"</td><td>"+$((Get-WsusServer).GetComputerTargetGroup($IDGroupes).GetParentTargetGroup().Name)+"</td><td>"+$((Get-WsusServer).GetComputerTargetGroup($IDGroupes).GetComputerTargets().Count)+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
#endregion

$BodyHtml += "<H2>Status des postes</H2>"

#region Erreur Mise à  jour

$BodyHtml += @("

    <H2>Mise à  jour en Erreur</H2>

    <table>
        <tr>
            <th>Mise à  jour</th>
            
        </tr>
        
    
")
    $count =0

    foreach ($errorup in $ErrorUpdates)  {
    
        $ErrorUpdatesName = $errorup
        $ErrorUpdatesName
        $count += 1
        write $count
        $BodyHtml += "<tr><td>"+$ErrorUpdatesName.Update+"</td></tr>"
        

    }

   
    $BodyHtml += "</table>"


#endregion

#region Error Postes

$BodyHtml += @("
    
   
    <H3> Postes en erreur</H3>
    <p>Il se peut que l'antivirus ne soit pas à jour et cela bloque la montée de version de l'OS.  En cas de problème, il faut se connecter sur le poste,
    essayer manuellement de mettre à jour le poste sinon il faut regarder les logs du windows update pour comprendre la source du problème</p>
    <table>
         <tr>
            <th>Nom</th>
            <th>IP</th>
            <th>OS</th>
            <th>Dernier Rapport</th>
         </tr>
    ") 
    $ErrorComputer = Get-WsusComputer -IncludedInstallationStates Failed
    
    
    
    foreach ( $errorco in $ErrorComputer )
    {
   write   $errorco.fulldomainname
        
        $temp +=   "<tr><td>"+$errorco.fulldomainname+"</td><td>"+$errorco.IPAddress+"</td><td>"+$errorco.osfamily+"</td><td>"+$errorco.lastsynctime+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
    
#endregion


#region Redémarrage Postes

$BodyHtml += @("
    
   
    <H3> Postes en attente de redémarrage</H3>
    <p> Il faut demander aux utilisateurs des postes de redémarrer leur ordinateur régulièrement pour que les mises à jour puissent s'installer</p>
    <table>
         <tr>
            <th>Nom</th>
            <th>IP</th>
            <th>OS</th>
            <th>Dernier Rapport</th>
         </tr>
    ") 
    
    
    foreach ( $pending in $PendingReboot )
    {
  
        
        $temp +=   "<tr><td>"+$pending.fulldomainname+"</td><td>"+$pending.IPAddress+"</td><td>"+$pending.osfamily+"</td><td>"+$pending.lastsynctime+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
    
#endregion



#region Download Postes

$BodyHtml += @("
   
   
    <H3> Mise à  jour téléchargée sur les postes</H3>
    <p> Les mises à jour ont été téléchargées sont en attente d'installation</p> 
    <table>
         <tr>
            <th>Nom</th>
            <th>IP</th>
            <th>OS</th>
            <th>Dernier Rapport</th>
         </tr>
    ") 
    
    
    foreach ( $downlo in $Dowloaded )
    {
  
        
        $temp +=   "<tr><td>"+$downlo.fulldomainname+"</td><td>"+$downlo.IPAddress+"</td><td>"+$downlo.osfamily+"</td><td>"+$downlo.lastsynctime+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
    
#endregion

 #region Installé Postes

$BodyHtml += @("
    <p> Les postes sont à jour </p>
   
    <H3> Postes avec les mises à  jour installé</H3>

    <table>
         <tr>
            <th>Nom</th>
            <th>IP</th>
            <th>OS</th>
            <th>Dernier Rapport</th>
         </tr>
    ") 
    
    
    foreach ( $SuccessCo in $SuccessComputer )
    {
  
        
        $temp +=   "<tr><td>"+$SuccessCo.fulldomainname+"</td><td>"+$SuccessCo.IPAddress+"</td><td>"+$SuccessCo.osfamily+"</td><td>"+$SuccessCo.lastsynctime+"</td></tr>"


    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
    
#endregion

#region Maj non installées

$BodyHtml += @("
    
   
    <H3> Mises à jour non installées sur les potes</H3>
    <p>Il se peut que le device doit redémarrer pour finaliser l'installation des mises à jour en cours avant de télécharger les nouvelles dont le device a besoin.
    <table>
         <tr>
            <th>Nom</th>
            <th>IP</th>
            <th>OS</th>
            <th>Dernier Rapport</th>
         </tr>
    ") 
  
    
    foreach ( $notinstall in $NotInstalled )
    {
  
        
        $temp +=   "<tr><td>"+$notinstall.fulldomainname+"</td><td>"+$notinstall.IPAddress+"</td><td>"+$notinstall.osfamily+"</td><td>"+$notinstall.lastsynctime+"</td></tr>"
       
    }


    $BodyHtml += $temp
    $temp = $null


       
    $BodyHtml += "</table>"
    
    
    
#endregion

$EndHtml += @("

    
    </body>
        </div>
             </div>
</html>


")

$HeadHTML+$BodyHtml+$EndHtml | Out-File $HtmlFile -Encoding utf8 -Force
$BodyHtml = $null












