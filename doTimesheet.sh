#!/bin/bash
cd "$(dirname "$0")"

inpPath=''
startCmd=''
dist=$(uname -a)
if [[ $dist  == *"WSL"* ]];
then
    echo 'Running in WSL'
    inpPath="/mnt/c/Users/figge/OneDrive - Epiroc/"
    startCmd='wslview'
else
    echo 'Running in ubuntu'
    inpPath='/home/ubuntu/OneDrive/'
    startCmd='xdg-open'
fi
echo "Creating timesheet..."
python3.11 ./timesheet.py "$inpPath"

# timesheet.py creates a 'timesheet.xlsx'
# Rename to reflect today's date
datetoday=$(date -I)
tdFileName='timesheet_'$datetoday'.xlsx'
mv timesheet.xlsx $tdFileName
cmdLine="${startCmd} ${tdFileName} &"
# echo $cmdLine
eval $cmdLine
