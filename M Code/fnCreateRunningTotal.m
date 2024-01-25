( #"Aggregation Column Name" as text, #"Target Table" as table, #"Target Column" as text) =>
let
    Source = #"Target Table",
    BuffValues = List.Buffer( Table.Column( #"Target Table", #"Target Column" ) ),
    RunningTotal =
      List.Generate (
        () => [ RunningTotal = BuffValues{0}, RowIndex = 0 ],
        each  [RowIndex] < List.Count(BuffValues),
        each  [ RunningTotal = List.Sum( { [RunningTotal] , BuffValues{[RowIndex] + 1} } ),
                RowIndex = [RowIndex] + 1 ],
        each  List.Max({1, [RunningTotal]}) ),
    #"Combined Table + RunningTotal" =
      Table.FromColumns(
        Table.ToColumns( #"Target Table" )
           & { Value.ReplaceType( RunningTotal, type {Int64.Type} ) } ,
        Table.ColumnNames( #"Target Table" ) & { #"Aggregation Column Name" } )
in
    #"Combined Table + RunningTotal"
