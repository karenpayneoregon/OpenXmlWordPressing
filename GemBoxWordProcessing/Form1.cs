using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GemBoxWordProcessing
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void cmdCreateNewEmptyDocument_Click(object sender, EventArgs e)
        {
            const string fileName = "NewEmptyDocument.docx";

            ((Button)sender).ImageKey = "(none)";

            var ops = new Operations();

            if (ops.CreateEmptyDocument(fileName))
            {
                ((Button)sender).ImageKey = "Success";
            }
            else
            {
                ((Button)sender).ImageKey = "Failed";
                MessageBox.Show($"Encountered problems creating '{fileName}'");
            }
        }

        private void cmdCreateNewDocumentWithSimpleParagraph_Click(object sender, EventArgs e)
        {
            const string fileName = "NewDocumentWithSingleParagraph.docx";

            ((Button)sender).ImageKey = "(none)";

            var ops = new Operations();

            if (ops.CreateDocumentWithSimpleParagraph(fileName))
            {
                ((Button)sender).ImageKey = "Success";
            }
            else
            {
                ((Button)sender).ImageKey = "Failed";
                MessageBox.Show($"Encountered problems creating '{fileName}'");
            }
        }

        private void cmdCreateNewDocumentWithUnoderedList_Click(object sender, EventArgs e)
        {
            const string fileName = "NewDocumentWithUnorderedList.docx";

            ((Button)sender).ImageKey = "(none)";

            var ops = new Operations();

            // this example differs from all the other code samples where 
            // all write operations are encapsulated into one class.
            if (ops.CreateDocumentWithUnoderedList(fileName))
            {
                ((Button)sender).ImageKey = "Success";
            }
            else
            {
                ((Button)sender).ImageKey = "Failed";
                MessageBox.Show($"Encountered problems creating '{fileName}'");
            }
        }

        private void cmdCreateNewDocumentWithMultipleParagraphAndImage_Click(object sender, EventArgs e)
        {
            const string fileName = "NewDocumentWithThreeParagraphWithImage.docx";

            ((Button)sender).ImageKey = "(none)";

            var ops = new Operations();

            if (ops.CreateDocumentWithMultipleParagraphAndImage(fileName))
            {
                ((Button)sender).ImageKey = "Success";
            }
            else
            {
                ((Button)sender).ImageKey = "Failed";
                MessageBox.Show($"Encountered problems creating '{fileName}'");
            }
        }
        private void cmdCreateNewDocumentWithTableFromArray_Click(object sender, EventArgs e)
        {
            const string fileName = "TableArray.docx";

            ((Button)sender).ImageKey = "(none)";

            var ops = new Operations();
            if (ops.CreateDocumentWithTableSimple(fileName))
            {
                ((Button)sender).ImageKey = "Success";
            }
            else
            {
                ((Button)sender).ImageKey = "Failed";
                MessageBox.Show($"Encountered problems creating '{fileName}'");
            }
        }
    }
}
