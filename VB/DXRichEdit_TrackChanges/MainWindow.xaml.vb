Imports System.Linq
Imports DevExpress.Xpf.Core
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace DXRichEdit_TrackChanges
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits ThemedWindow

		Public Sub New()
			InitializeComponent()

			richEditControl1.LoadDocument("DocumentWithRevisions.docx")
			richEditControl1.AnnotationOptions.VisibleAuthors.Remove("Michael Suyama")

			AcceptAndRejectRevisions()
			richEditControl1.Document.SaveDocument("DocumentWithRevisions.docx", DocumentFormat.OpenXml)


		End Sub

		Private Sub AcceptAndRejectRevisions()
			Dim documentRevisions As RevisionCollection = richEditControl1.Document.Revisions

			'Reject all revisions in the firts page's header:
			Dim header As SubDocument = richEditControl1.Document.Sections(0).BeginUpdateHeader(HeaderFooterType.First)
			documentRevisions.RejectAll(header)
			richEditControl1.Document.Sections(0).EndUpdateHeader(header)

			'Reject all revisions from the specific author on the first section:
			Dim sectionRevisions = documentRevisions.Get(richEditControl1.Document.Sections(0).Range).Where(Function(x) x.Author = "Janet Leverling")

			For Each revision As Revision In sectionRevisions
				revision.Reject()
			Next revision

			'Accept all format changes:
			documentRevisions.AcceptAll(Function(x) x.Type = RevisionType.CharacterPropertyChanged OrElse x.Type = RevisionType.ParagraphPropertyChanged OrElse x.Type = RevisionType.SectionPropertyChanged)
		End Sub

		Private Sub RichEditControl_TrackedMovesConflict(ByVal sender As Object, ByVal e As TrackedMovesConflictEventArgs)
			'Compare the length of the original and new location ranges
			'Keep text from the location whose range is the smallest
			e.ResolveMode = If(e.OriginalLocationRange.Length <= e.NewLocationRange.Length, TrackedMovesConflictResolveMode.KeepOriginalLocationText, TrackedMovesConflictResolveMode.KeepNewLocationText)
		End Sub
	End Class
End Namespace
