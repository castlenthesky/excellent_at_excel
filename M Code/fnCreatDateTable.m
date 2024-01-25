(#"Start Date" as date, #"End Date" as date) as any =>
let
    date_list = List.Dates(#"Start Date", Duration.Days(#"End Date" - #"Start Date")+1, #duration(1,0,0,0) ),
    date_table = Table.FromList(date_list, Splitter.SplitByNothing(), {"Date"}, null),
    #"Changed Type" = Table.TransformColumnTypes(date_table,{{"Date", type date}}),
    #"Inserted Year" = Table.AddColumn(#"Changed Type", "Year", each Date.Year([Date]), Int64.Type),
    #"Inserted Quarter of Year" = Table.AddColumn(#"Inserted Year", "Quarter of Year", each Date.QuarterOfYear([Date]), Int64.Type),
    #"Inserted Month of Year" = Table.AddColumn(#"Inserted Quarter of Year", "Month of Year", each Date.Month([Date]), Int64.Type),
    #"Inserted Day of Month" = Table.AddColumn(#"Inserted Month of Year", "Day of Month", each Date.Day([Date]), Int64.Type),
    #"Inserted Day of Quarter" = Table.AddColumn(#"Inserted Day of Month", "Day of Quarter", each Duration.Days([Date] - Date.StartOfQuarter([Date])) + 1, Int64.Type),
    #"Inserted Day of Year" = Table.AddColumn(#"Inserted Day of Quarter", "Day of Year", each Date.DayOfYear([Date]), Int64.Type),
    #"Inserted Week of Year" = Table.AddColumn(#"Inserted Day of Year", "Week of Year", each Number.RoundUp([Day of Year]/7), Int64.Type),
    #"Inserted Week of Month" = Table.AddColumn(#"Inserted Week of Year", "Week of Month", each Date.WeekOfMonth([Date]), Int64.Type),
    #"Inserted Day of Week" = Table.AddColumn(#"Inserted Week of Month", "Day of Week", each Date.DayOfWeek([Date]), Int64.Type),
    #"Inserted Days in Month" = Table.AddColumn(#"Inserted Day of Week", "Days in Month", each Date.DaysInMonth([Date]), Int64.Type),
    #"Inserted Week of Quarter" = Table.AddColumn(#"Inserted Days in Month", "Week of Quarter", each Number.RoundUp([Day of Quarter]/7), Int64.Type),
    #"Inserted Is Weekday" = Table.AddColumn(#"Inserted Week of Quarter", "Is Weekday", each if [Day of Week] > 0 and [Day of Week] < 6 then 1 else 0, Int64.Type),
    #"Inserted Holiday Name" = Table.AddColumn(#"Inserted Is Weekday", "Holiday Name", each if
        // New Year's Eve -- December 31st
        [Month of Year] = 12 and [Day of Month] = 31 then "New Year's Eve"
        else if [Month of Year] = 12 and [Day of Month] + 1 = 31 and [Day of Week] = 5 then "New Year's Eve Observed"
        else if [Month of Year] = 12 and [Day of Month] + 2 = 31 and [Day of Week] = 5 then "New Year's Eve Observed"
        // New Year's Day -- January 1st
        else if [Day of Year] = 1 then "New Year's Day"
        else if [Day of Year] = 2 and [Day of Week] = 1 then "New Year's Observed"
        else if [Day of Year] = 3 and [Day of Week] = 1 then "New Year's Observed"
        // Martin Luther King Day -- 3rd Monday in January
        else if [Month of Year] = 1 and [Week of Month] = 3 and [Day of Week] = 1 then "Martin Luther King Day"
        // President's Day -- 3rd Monday in February
        else if [Month of Year] = 2 and [Week of Month] = 3 and [Day of Week] = 1 then "President's Day"
        // Memorial Day -- Last Monday in May
        else if [Month of Year] = 5 and [Day of Week] = 1 and [Days in Month] - [Day of Month] <= 6 then "Memorial Day"
        // Independence Day -- July 4th
        else if [Month of Year] = 7 and [Day of Month] = 4 then "Independance Day"
        else if [Month of Year] = 7 and [Day of Month] = 3 and [Day of Week] = 5 then "Independance Day Observed"
        else if [Month of Year] = 7 and [Day of Month] = 5 and [Day of Week] = 1 then "Independance Day Observed"
        // Labour Day -- First Monday in September
        else if [Month of Year] = 9 and [Day of Week] = 1 and [Day of Month] <= 7 then "Labour Day"
        // Juneteenth -- June 19th
        else if [Month of Year] = 9 and [Day of Month] = 19 then "Juneteenth"
        else if [Month of Year] = 9 and [Day of Month] = 18 and [Day of Week] = 5 then "Juneteenth Observed"
        else if [Month of Year] = 9 and [Day of Month] = 20 and [Day of Week] = 1 then "Juneteenth Observed"
        // Halloween -- October 31st
        else if [Month of Year] = 10 and [Day of Month] = 31 then "Halloween"
        // Veterans' Day -- November 11th
        else if [Month of Year] = 11 and [Day of Month] = 11 then "Veteran's Day"
        else if [Month of Year] = 11 and [Day of Week] = 5 and [Day of Month] = 10 then "Veteran's Day Observed"
        else if [Month of Year] = 11 and [Day of Week] = 1 and [Day of Month] = 12 then "Veteran's Day Observed"
        // Christmas Eve -- December 24th
        else if [Month of Year] = 12 and [Day of Month] = 24 then "Christmas Eve"
        else if [Month of Year] = 12 and [Day of Week] = 5 and [Day of Month] = 23 then "Christmas Eve Observed"
        else if [Month of Year] = 12 and [Day of Week] = 5 and [Day of Month] = 22 then "Christmas Eve Observed"
        // Christmas Day -- December 25th
        else if [Month of Year] = 12 and [Day of Month] = 25 then "Christmas Day"
        else if [Month of Year] = 12 and [Day of Week] = 1 and [Day of Month] = 26 then "Christmas Day Observed"
        else if [Month of Year] = 12 and [Day of Week] = 1 and [Day of Month] = 27 then "Christmas Day Observed"
        else null, type text),
    #"Inserted Is Workday" = Table.AddColumn(#"Inserted Holiday Name", "Is Workday", each if [Is Weekday] = 1 and [Holiday Name] = null then 1 else 0, Int64.Type),
    #"Inserted Workdays in Month" = Table.ExpandTableColumn(
        Table.NestedJoin(
            #"Inserted Is Workday",
            {"Year", "Month of Year"},
            Table.Group(#"Inserted Is Workday", {"Year", "Month of Year"}, {{"Workdays in Month", each List.Sum([Is Workday]), Int64.Type}}),
            {"Year", "Month of Year"},
            "Aggregated Months", JoinKind.LeftOuter),
        "Aggregated Months",
        {"Workdays in Month"},
        {"Workdays in Month"}),
    #"Inserted UTC Diff" = Table.AddColumn(#"Inserted Workdays in Month", "UTC Diff", each DateTimeZone.ZoneHours(DateTimeZone.ToLocal(DateTime.AddZone(DateTime.From([Date]),0))),Int64.Type),
    #"Inserted Workday of Month" = Table.Combine(Table.Group(
        #"Inserted UTC Diff",
        {"Year", "Month of Year"},
        {{"Details", each fnCreateRunningTotal("Workday of Month", _, "Is Workday"), Int64.Type}})[Details]),
    #"Inserted Workday of Quarter" = Table.Combine(Table.Group(
        #"Inserted Workday of Month",
        {"Year", "Quarter of Year"},
        {{"Details", each fnCreateRunningTotal("Workday of Quarter", _, "Is Workday"), Int64.Type}})[Details]),
    #"Inserted Workday of Year" = Table.Combine(Table.Group(
        #"Inserted Workday of Quarter",
        {"Year"},
        {{"Details", each fnCreateRunningTotal("Workday of Year", _, "Is Workday"), Int64.Type}})[Details]),
    #"Ordered Columns" = Table.ReorderColumns(#"Inserted Workday of Year",{"Date", "Year", "Quarter of Year", "Month of Year", "Day of Month", "Day of Week", "Day of Quarter", "Day of Year", "Week of Month", "Week of Quarter", "Week of Year", "Days in Month", "Is Weekday", "Is Workday", "Holiday Name", "Workdays in Month", "Workday of Month", "Workday of Quarter", "Workday of Year", "UTC Diff"})
in
    #"Ordered Columns"
