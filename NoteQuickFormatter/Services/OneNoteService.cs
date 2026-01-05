using Microsoft.Office.Interop.OneNote;
using NoteQuickFormatter.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Collections.Specialized.BitVector32;

namespace NoteQuickFormatter
{
    public class OneNoteService
    {
        private readonly Application _oneNote = new Application();
        private XDocument _doc;
        private XNamespace _ns;
        private Notebook[] _notebooks;
        private XElement _currentlyViewedNotebook;
        private string _firstNewlyCreatedPageID = "";
        public XElement CurrentlyViewedNotebook => _currentlyViewedNotebook;
        public string[] NotebookNames => _notebooks.Select(n => n.Name).ToArray();

        public OneNoteService()
        {
            RefreshHierarchy();
        }

        // notebook -> section -> page
        public void RefreshHierarchy(HierarchyScope scope = HierarchyScope.hsSections)
        {
            string xml;
            try
            {
                _oneNote.GetHierarchy(null, scope, out xml);
            }
            catch (Exception ex)
            {
                throw new Exception($"Exception from _oneNote.GetHierarchy Error: {ex}");
            }
            _doc = XDocument.Parse(xml);
            _ns = _doc.Root.Name.Namespace;
            var notebooks = _doc.Descendants(_ns + "Notebook").ToArray();
            _notebooks = notebooks.Select(n =>
            {
                string name = n.Attribute("nickname") == null ? n.Attribute("name").Value : n.Attribute("nickname").Value;
                bool isCurrentlyViewed = n.Attribute("isCurrentlyViewed") != null;
                var nb = new Notebook()
                {
                    ID = n.Attribute("ID").Value,
                    Name = name,
                    IsCurrentlyViewed = isCurrentlyViewed,
                    Path = n.Attribute("path").Value
                };
                if (isCurrentlyViewed)
                { _currentlyViewedNotebook = n; }
                return nb;
            }).ToArray();
        }

        private XElement GetPageTitleElement(string title)
        {
            return new XElement(_ns + "Title",
                        new XElement(_ns + "OE",
                        new XAttribute("style", "font-family:'Bradley Hand ITC';font-size:20.0pt"),
                        new XElement(_ns + "T", title)));
        }
        private XElement GetTagDef()
        {
            XElement tagDef = new XElement(_ns + "TagDef",
                                new XAttribute("index", 0),
                                new XAttribute("type", 0),
                                new XAttribute("symbol", 3),
                                new XAttribute("name", "待辦事項"));
            return tagDef;
        }
        private XElement GetFormattedPageContentElement(string title)
        {
            // 位置是為了符合手動新增時，輸入完標題後按下enter會跳至的位置
            XElement position = new XElement(_ns + "Position",
                        new XAttribute("x", "36.0"),
                        new XAttribute("y", "86.4000015258789"),
                        new XAttribute("z", "0"));
            // 尺寸是隨便設定的
            XElement size = new XElement(_ns + "Size",
                new XAttribute("width", "120"),
                new XAttribute("height", "120")
                );
            XElement oeChildren = new XElement(_ns + "OEChildren");
            XElement outline = new XElement(_ns + "Outline", position, size, oeChildren);

            DateTime startDate = DateTimeHelper.ConvertStringToDateTime(title);
            for (int i = 0; i < 5; i++)
            {
                XElement date = new XElement(_ns + "T", startDate.AddDays(i).ToString("M/d"));
                XElement toDo = new XElement(_ns + "OE",
                    new XElement(_ns + "Tag",
                    new XAttribute("index", 0)),
                    new XElement(_ns + "T"));
                XElement newLine = new XElement(_ns + "OE", new XElement(_ns + "T"));
                XElement content = new XElement(_ns + "OEChildren", toDo, newLine);
                XElement oe = new XElement(_ns + "OE",
                    date,
                    content);
                oeChildren.Add(oe);
            }
            return outline;
        }
        public void CreateNewSection(bool addOneMonth = false)
        {
            string[] abbreviations = new DateTimeHelper().AbbreviatedMonthNames;
            int month = DateTime.Now.Month;
            string sectionName = addOneMonth ? abbreviations[month % 12] : abbreviations[month - 1];
            CreateNewSectionElement(sectionName, _currentlyViewedNotebook, addOneMonth);
        }
        public void CreateNewSection(string name, string notebookName, bool addOneMonth)
        {
            var notebook = GetNotebookByNickname(notebookName);
            CreateNewSectionElement(name, notebook, addOneMonth);
        }

        public XElement GetNotebookByNickname(string nickname)
        {
            XElement notebook = _doc.Descendants(_ns + "Notebook").FirstOrDefault(nb => nb.Attribute("nickname").Value == nickname);
            return notebook;
        }
        public void CreateNewSectionElement(string name, XElement notebook, bool addOneMonth)
        {
            RefreshHierarchy();
            DateTime day = addOneMonth ? DateTime.Today.AddMonths(1) : DateTime.Now;
            var pageTitles = DateTimeHelper.GetWeekdayRanges(day.Year, day.Month);
            XElement section = new XElement(_ns + "Section",
                new XAttribute("name", name));

            if (notebook.Elements(_ns + "Section").Contains(section, new ElementNameComparer()))
            {
                throw new Exception("Already has a section with the same name.");
            }
            for (int i = 0; i < pageTitles.Count; i++)
            {
                section.Add(new XElement(_ns + "Page"));
            }
            notebook.Elements(_ns + "Section").LastOrDefault().AddAfterSelf(section);
            try
            {
                _oneNote.UpdateHierarchy(notebook.ToString());
                RefreshHierarchy(HierarchyScope.hsPages);
                notebook = GetNotebookByNickname(notebook.Attribute("nickname").Value);
                AddPageContent(notebook, name, pageTitles);
                string id = notebook.Elements(_ns + "Section")
                                    .FirstOrDefault(s => s.Attribute("name").Value == name)
                                    .Attribute("ID").Value;
                _oneNote.NavigateTo(id);
            }
            /* TODO */
            // 如果沒有先取得hierarchy，會跳出沒有ns的錯誤
            // 或許會找不到ID?
            catch (Exception ex)
            { System.Windows.Forms.MessageBox.Show(ex.Message); }
        }
        public void AddPageContent(XElement notebook, string sectionName, List<string> pageTitles)
        {
            try
            {
                XElement section = notebook.Elements(_ns + "Section")
                                           .FirstOrDefault(s => s.Attribute("name").Value == sectionName);
                var pages = section.Descendants(_ns + "Page")
                                   .ToArray();
                for (int i = 0; i < pageTitles.Count; i++)
                {
                    pages[i].Add(GetTagDef());
                    pages[i].Add(GetPageTitleElement(pageTitles[i]));
                    pages[i].Add(GetFormattedPageContentElement(pageTitles[i].Split('-')[0]));
                    _oneNote.UpdatePageContent(pages[i].ToString());
                    RefreshHierarchy();
                }
            }
            /* TODO */
            catch (Exception ex)
            { }
        }
        
        public string GetPageContent(string id)
        {
            _oneNote.GetPageContent(id, out string xml);
            return xml;
        }
    }

    public enum OneNoteSectionColor
    {

    }
}
