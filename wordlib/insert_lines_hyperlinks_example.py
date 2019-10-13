def insert_lines(d, lines):
    """ d: word document object
    lines: tuple (int, str, str) of 
    level, text to display, href. """
    pgs = d.Paragraphs
    p = pgs(pgs.Count)
    r = p.Range

    w = d.Application
    w.ScreenUpdating = False
    apply_list_template(lt, r)
   
    try:
     for i, (l, s, h) in enumerate(lines,1):
         p.Format.TabStops.Add(inches_to_points(6), c.wdAlignTabLeft, c.wdTabLeaderSpaces)
         l+=1
         d.Hyperlinks.Add(r, h, "", "Open in Webbrowser", "<link>")
         r.InsertBefore(s + "\t")
         r.ListFormat.ListLevelNumber = l
         
         if not i % 10:
             print("\r%d/%d         " % (i, len(lines)),end="")
         pgs.Add()
         p=pgs(pgs.Count)
         r=p.Range
    finally:
        w.ScreenUpdating=True
        
        
def insert_lines2(lines):
 r=pgs(pgs.Count).Range
 for lvl, txt, href in lines:
    p=pgs(pgs.Count)
    p.Format.LineSpacingRule = c.wdLineSpaceSingle
    p.Format.SpaceBeforeAuto = False
    p.Format.SpaceAfterAuto = False
    p.TabStops.Add(inches_to_points(6), c.wdAlignTabLeft, c.wdTabLeaderDots)
    r.Text = txt
    r.InsertAfter("\t")
    r.MoveStart(c.wdCharacter, len(txt)+2)
    d.Hyperlinks.Add(r, href, "", "Open Webbrowser", "Link")
    r.InsertAfter("\r")
    r.MoveStart(c.wdCharacter, len(href) + 10)
    

    
    
    
        
