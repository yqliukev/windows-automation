import win32com.client as win32
from time import sleep

# adds paragraphs from one file into new generated file
def substitution(paragraphs, new_path, base_path, company, position, skills):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True

    # First copy title and intro format from base file
    base_file = word.Documents.Open(FileName = base_path)
    rng = base_file.Range(Start = base_file.Paragraphs(1).Range.Start, End = base_file.Paragraphs(4).Range.End)
    rng.Select()
    sele = word.Selection
    sele.Copy()
    
    # Paste title into new file
    new_file = word.Documents.Add()
    new_file.Content.PasteAndFormat(Type = 16)
    
    # Expand range to cover whole document, and then search for needed paragraphs, paste into new document
    for paragraph in paragraphs:
        
        # search for needed paragraphs
        rng = base_file.Content
        rng.Select()
        fnd = sele.Find
        fnd.Text = paragraph
        fnd.Execute()
        sele.Expand(Unit = 4)
        sele.Copy()
        sele.Collapse()
        
        # paste into end of new doc
        rng = new_file.Content
        rng.Collapse(Direction = 0)
        rng.PasteAndFormat(Type = 16)
        
        # delete beginning category word with find
        rng = new_file.Content
        rng.Select()
        sele = word.Selection
        fnd = sele.Find
        fnd.Text = paragraph
        fnd.Execute()
        sele.Delete()

    # Copy and paste sendoff

    par_count = base_file.Paragraphs.Count
    rng = base_file.Range(Start = base_file.Paragraphs(par_count - 3).Range.Start, End = base_file.Paragraphs.Last.Range.End)
    rng.Select()
    sele = word.Selection
    sele.Copy()
    
    rng = new_file.Content
    rng.Collapse(Direction = 0)
    rng.PasteAndFormat(Type = 16)


    # find and replace company name
    rng = new_file.Content
    rng.Find.Execute(FindText = "[Company]", ReplaceWith = company, Replace = 2)

    # find and replace skills
    skill_str = skill_cat(skills)
    rng = new_file.Content
    rng.Find.Execute(FindText = "[skills]", ReplaceWith = skill_str, Replace = 2)

    # find and replace position name
    rng = new_file.Content
    rng.Find.Execute(FindText = "[position]", ReplaceWith = position, Replace = 2)


    # Save Newfile name
    new_file.SaveAs2(FileName = new_path)

    sleep(1)

def skill_cat(skills):
    result = ""
    for i in range(len(skills)):
        if i == len(skills) - 1:
            result = result + " and " + skills[i]
            return result
        else :
            result = result + " " + skills[i]


        

substitution(["_programming_"], "D:\Documents\Job\Cover Letters\Test Cover Letter", "D:\Documents\Job\Cover Letters\Base Cover Letter.docx", "Bicdroid Inc.", "QA", ["computer proficiency", "communication"])

