def insert_initial_date(cols=3):
    s = w.Selection
    s.InsertRowsBelow(1)
    w.Selection.Range.Text = "Initial"
    r = w.Selection.Range
    r.ParagraphFormat.Alignment = c.wdAlignParagraphLeft
    merge(r)
    w.Selection.InsertRowsBelow(1)
    w.Selection.Range.Text = "Date"
    r = w.Selection.Range
    r.ParagraphFormat.Alignment = c.wdAlignParagraphLeft
    merge(r)
 
insert2 = hide_alerts(insert_initial_date)

 