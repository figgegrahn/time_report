echo "Creating timesheet..."
python3 ./timesheet.py
# timesheet.py creates a 'timesheet.xlsx'
# Rename to reflect today's date
datetoday=$(date -I)
tdFileName='timesheet_'$datetoday'.xlsx'
mv timesheet.xlsx $tdFileName
xdg-open $tdFileName &
