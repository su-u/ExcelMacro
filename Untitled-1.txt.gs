

function NameSchedule() {
    const COLUMN_POSITION = 4
    const ROW_POSITION = 4
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var name_schedule_sheet = spreadsheet.getSheets()[0];
    var schedule_sheet      = spreadsheet.getSheets()[2];
    var management_sheet    = spreadsheet.getSheets()[3];

    const num_of_day = 2;
    const num_of_oneday = 2;
    const num_of_times = ((num_of_day * num_of_oneday));
    const num_of_people = 20;
    const time_per_people = 5;
    
    var rand = 0; 
    var average = 1;
    var people_count = 0;
    
    var i = 0;
    var j = 0;
    name_schedule_sheet.getRange(1 ,1).setValue(schedule_sheet.getRange(1,1).getValue());
    for (var i = COLUMN_POSITION; i < (num_of_times + COLUMN_POSITION); i++) {
        j = 0;
        do {
            rand = Math.floor(Math.random() * num_of_people);
            var number_of = name_schedule_sheet.getRange(rand + ROW_POSITION,1).getValue();
            if (schedule_sheet.getRange(rand + ROW_POSITION,i).getValue() != "1" && number_of < average) {
                name_schedule_sheet.getRange(rand + ROW_POSITION ,i).setValue(1);
                number_of++;
                name_schedule_sheet.getRange(rand + ROW_POSITION ,1).setValue(number_of);
                j++;
                people_count++;
              　if(people_count === num_of_people){
                  people_count = 0;
                  average++;
              　}
            }
        } while (j < time_per_people);
    }
}

