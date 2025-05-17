#Requires -Version 7.3
#Requires -RunAsAdministrator

param()

$InformationPreference = 'Continue'



# Setup target-apps
Write-Information 'Setup: git...'
git config --system core.longpaths true

git config --system init.defaultbranch 'main'

git config --global user.name 'joykler'
git config --global user.email 'joykler@gmail.com'



# Done!
Write-Information "`n[DONE]"
