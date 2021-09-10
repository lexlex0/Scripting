##################################################################################### 
#####       	                 Script O365 Migration Licences					#####
#####									Office 365								#####
#####     						    Alexandre Clerbois							#####
#####							      Computerland								#####
#####################################################################################
#install-Module ExchangeOnlineManagement -Force
#Install-Module MSOnline -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Try
{
		Connect-MsolService
	
		Get-MsolUser -EnabledFilter EnabledOnly -MaxResult 50000 | Select-Object UserPrincipalName > User.txt
		$msgBoxInput = [System.Windows.MessageBox]::Show('Voulez-vous generez un document avec la liste des utilisateurs actuels pour la Migration des utilisateurs des licences ?', 'Ouvrir Fichier Genere', 'YesNo')
		If ($msgBoxInput -eq 'Yes')
		{
			Get-Content .\User.txt | Select-Object -Skip 1 | Out-File .\User1.txt
			Get-Content .\User1.txt | Select-Object -Skip 1 | Out-File .\User2.txt
			Get-Content .\User2.txt | Select-Object -Skip 1 | Out-File .\User.txt
			Remove-Item .\User1.txt
			Remove-Item .\User2.txt
			.\User.txt
			[System.Windows.MessageBox]::Show("Un Document txt a ete lancer,`n Supprimer les utilisateurs dont vous ne souhaitez pas Migration licences`n Une fois termine sauvegarder le txt et fermer le.`nAppuyer sur OK quand vous aurez terminer", 'Information Importante (del Licences)')
			$File = Get-Content .\User.txt
		}
		else
		{
			$msgBoxInput1 = [System.Windows.MessageBox]::Show('Voulez-vous ouvrir une liste specifique (TXT) ? `n La liste doit seulement avoir les adresses des utilisateurs pour etre pris en compte !`n`n Sinon un fichier vide sera ouvert', 'Ouvrir Fichier (del Licences)', 'YesNo')
			If ($msgBoxInput1 -eq 'Yes')
			{
				$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
					InitialDirectory = [Environment]::GetFolderPath('Desktop')
					Filter		     = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
					Title		     = "Ouvrir Liste Suppression Licences "
				}
				$null = $FileBrowser.ShowDialog()
				$File = Get-Content $FileBrowser.FileName
			}
			Else
			{
				Write-Output "" > .\User.txt
				.\User.txt
				[System.Windows.MessageBox]::Show("Un Document txt vide a ete lancer,`n Ajouter la liste des utilisateurs que vous souhaitez Migrations des licences.`nUne fois termine sauvegarder le txt et fermer le.`nAppuyer sur OK quand vous aurez terminer", 'Information Importante')
				$File = Get-Content .\User.txt
			}
		}
		[System.Windows.MessageBox]::Show("Merci de selectionner la Licences a SUPPRIMER", 'Information Importante')
		$ChoixLic = Get-MsolAccountSku | Out-GridView -OutputMode single -Title "Selectionnez une licences A SUPPRIMER"
		[System.Windows.MessageBox]::Show("Merci de selectionner la Licences a Remplacer", 'Information Importante')
		$ChoixLicNew = Get-MsolAccountSku | Out-GridView -OutputMode Single -Title "Selectionnez une licences A REMPLACER"
		$Error.Clear()
		$chooseDel = $ChoixLic.AccountSkuId
		$chooseNew = $ChoixLicNew.AccountSkuId
		Clear-Host
		$iLine = 0
foreach ($line in $File)
{
	if ($line -match $regex)
	{
		Set-MsolUser -UserPrincipalName $line -UsageLocation BE
		$Perc = ($iLine/$File.Count)/100
		$iLine++
		Try
		{
			Set-MsolUserLicense -UserPrincipalName $line -RemoveLicenses $chooseDel -ErrorAction Stop
			Set-MsolUserLicense -UserPrincipalName $line -AddLicenses $chooseNew -ErrorAction Stop
			Write-Progress -Activity "Processed mailbox Delete license $line "
		}
		Catch
		{
			Write-Output "ERREUR pour $Line"
			$tableErreur += @([pscustomobject]@{ Email = $Line; Error = $Error[0].Exception.Message })
		}
	}
}
	if ($Error.count -ne 0)
	{
		$MSGError1 = [System.Windows.MessageBox]::Show('Au moins une erreur a ete detecte, voulez vous afficher la liste des utilisateurs en erreur ?', 'Erreur Detecte', 'YesNo')
		if ($MSGError1 -eq "Yes")
		{
			Write-Output "`n`n`n`n"
			Write-Output $tableErreur
			$MSGError2 = [System.Windows.MessageBox]::Show('Voulez-vous exportez la liste d erreur ?', 'Erreur detect', 'YesNo')
			if ($MSGError2 -eq 'Yes')
			{
				$Sorti = Select-Folder "Selectionner la destination du fichier d'erreur"
				$tableErreur | Export-Csv $Sorti"\ErrorDelLic$DateNow.csv"
			}
			Pause
		}
	}
}
Catch
{
	Write-Host "Erreur de connexion MSOnline"
}