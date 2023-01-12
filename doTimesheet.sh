dist=$(uname -a)
inpPath=''
startCmd=''

if [[ $dist =~ 'WSL' ]]
then
    echo 'Running in WSL'
    inpPath='/mnt/c/Users/tmgfgn/Downloads/'
    startCmd='cmd.exe /C start'
else
    echo 'Running in linux'
    inpPath='/home/ubuntu/OneDrive/'
    startCmd='xdg-open'
fi
echo "Creating timesheet..."
python3 ./timesheet.py $inpPath

# timesheet.py creates a 'timesheet.xlsx'
# Rename to reflect today's date
datetoday=$(date -I)
tdFileName='timesheet_'$datetoday'.xlsx'
mv timesheet.xlsx $tdFileName
eval $startCmd $tdFileName &
