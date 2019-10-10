using System.Linq;
using DevExpress.Xpf.Core;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace DXRichEdit_TrackChanges
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ThemedWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            richEditControl1.LoadDocument("DocumentWithRevisions.docx");
            richEditControl1.AnnotationOptions.VisibleAuthors.Remove("Michael Suyama");

            AcceptAndRejectRevisions();
            richEditControl1.Document.SaveDocument("DocumentWithRevisions.docx", DocumentFormat.OpenXml);


        }

        private void AcceptAndRejectRevisions()
        {
            RevisionCollection documentRevisions = richEditControl1.Document.Revisions;

            //Reject all revisions in the firts page's header:
            SubDocument header = richEditControl1.Document.Sections[0].BeginUpdateHeader(HeaderFooterType.First);
            documentRevisions.RejectAll(header);
            richEditControl1.Document.Sections[0].EndUpdateHeader(header);

            //Reject all revisions from the specific author on the first section:
            var sectionRevisions = documentRevisions.Get(richEditControl1.Document.Sections[0].Range).Where(x => x.Author == "Janet Leverling");

            foreach (Revision revision in sectionRevisions)
                revision.Reject();

            //Accept all format changes:
            documentRevisions.AcceptAll(x => x.Type == RevisionType.CharacterPropertyChanged || x.Type == RevisionType.ParagraphPropertyChanged || x.Type == RevisionType.SectionPropertyChanged);
        }

        private void RichEditControl_TrackedMovesConflict(object sender, TrackedMovesConflictEventArgs e)
        {
            //Compare the length of the original and new location ranges
            //Keep text from the location whose range is the smallest
            e.ResolveMode = (e.OriginalLocationRange.Length <= e.NewLocationRange.Length) ? TrackedMovesConflictResolveMode.KeepOriginalLocationText : TrackedMovesConflictResolveMode.KeepNewLocationText;
        }
    }
}
