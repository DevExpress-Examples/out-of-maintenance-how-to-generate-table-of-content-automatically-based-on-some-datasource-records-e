using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Office;
using DevExpress.XtraRichEdit.API.Native;

namespace WindowsFormsApplication1 {
    public partial class Form1 : RibbonForm {
        public Form1() {
            InitializeComponent();
        }

        DataTable someDT = new DataTable();
        private void Form1_Load(object sender, EventArgs e) {
            string loremIpsumString =  "Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, sit amet commodo magna eros quis urna.\r\n" +
                                       "Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus.\r\n" +
                                       "Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra nonummy pede. Mauris et orci.";


            someDT.Columns.Add("Header", typeof(string));
            someDT.Columns.Add("Content", typeof(string));

            for(int i = 0; i < 10; i++) {
                someDT.Rows.Add(new object[] { "Header " + i.ToString(), loremIpsumString });
            }
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            foreach(DataRow dataRow in someDT.Rows) {
                Paragraph headerParagraph = richEditControl1.Document.Paragraphs.Append();
                SetParagraphStyleSettings(headerParagraph, 1);
                DocumentRange headerRange = richEditControl1.Document.AppendText(dataRow["Header"].ToString());
                Paragraph singleParagraph = richEditControl1.Document.Paragraphs.Append();
                SetParagraphStyleSettings(singleParagraph, 0);
                ChangeHeaderCharacterProeprties(headerRange);
                richEditControl1.Document.AppendText(dataRow["Content"].ToString());
                richEditControl1.Document.AppendText(Characters.PageBreak.ToString());                
            }

            InsertTOC("\\h");
        }

        void InsertTOC(string switches) {
            Field field = richEditControl1.Document.Fields.Create(richEditControl1.Document.Range.Start, "TOC " + switches);
            CharacterProperties cp = richEditControl1.Document.BeginUpdateCharacters(field.Range);
            cp.Bold = false;
            cp.FontSize = 12;
            cp.ForeColor = Color.Blue;
            richEditControl1.Document.EndUpdateCharacters(cp);
            richEditControl1.Document.InsertSection(field.Range.End);
            field.Update();        
        }

        int levelIndent = 1;
        private void SetParagraphStyleSettings(Paragraph currentParagraph, int level) {
            if(level > 0) {
                string styleName = "Paragraph Level " + levelIndent.ToString();
                ParagraphStyle paragraphStyle = richEditControl1.Document.ParagraphStyles[styleName];

                if(paragraphStyle == null) {
                    paragraphStyle = richEditControl1.Document.ParagraphStyles.CreateNew();
                    paragraphStyle.Name = styleName;
                    paragraphStyle.Parent = richEditControl1.Document.ParagraphStyles["Normal"];
                    paragraphStyle.OutlineLevel = levelIndent;
                    richEditControl1.Document.ParagraphStyles.Add(paragraphStyle);
                }
                currentParagraph.Style = paragraphStyle;
            }
            else {
                currentParagraph.Style = richEditControl1.Document.ParagraphStyles["Normal"];            
            }
        }

        private void ChangeHeaderCharacterProeprties(DocumentRange headerRange) {
            CharacterProperties cp = richEditControl1.Document.BeginUpdateCharacters(headerRange);
            cp.Bold = true;
            cp.ForeColor = Color.Blue;
            richEditControl1.Document.EndUpdateCharacters(cp);
        }
    }
}
