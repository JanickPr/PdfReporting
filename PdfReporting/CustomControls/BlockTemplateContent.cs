using System.Windows;
using System.Windows.Documents;

namespace WpfPdfReporting.CustomControls
{
    public class BlockTemplateContent : Section
    {
        private static readonly DependencyProperty TemplateProperty = DependencyProperty.Register("Template", typeof(DataTemplate), typeof(BlockTemplateContent), new PropertyMetadata(OnTemplateChanged));

        public DataTemplate Template
        {
            get => (DataTemplate)GetValue(TemplateProperty);
            set => SetValue(TemplateProperty, value);
        }

        public BlockTemplateContent()
        {
            Helpers.FixupDataContext(this);
            Loaded += this.BlockTemplateContent_Loaded;
        }

        private void BlockTemplateContent_Loaded(object sender, RoutedEventArgs e)
        {
            GenerateContent(this.Template);
        }

        private void GenerateContent(DataTemplate template)
        {
            this.Blocks.Clear();
            if (template != null)
            {
                FrameworkContentElement element = Helpers.LoadDataTemplate(template);
                this.Blocks.Add((Block)element);
            }
        }

        private void OnTemplateChanged(DataTemplate dataTemplate)
        {
            GenerateContent(dataTemplate);
        }

        private static void OnTemplateChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            ((BlockTemplateContent)d).OnTemplateChanged((DataTemplate)e.NewValue);
        }
    }
}
