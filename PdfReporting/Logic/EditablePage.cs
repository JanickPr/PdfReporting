using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace PdfReporting.Logic
{
    public class EditablePage : ContainerVisual
    {
        public void AddVisualAtTopOfPage(Visual visual)
        {
            ContainerVisual containerVisual = GetContainerVisualWith(0);
            containerVisual.Children.Add(visual);
            this.Children.Add(containerVisual);
        }

        public void AddVisualAt(double offsetY, Visual visual)
        {
            if (visual == null) return;
            ContainerVisual containerVisual = GetContainerVisualWith(offsetY);
            containerVisual.Children.Add(visual);
            this.Children.Add(containerVisual);
        }

        private ContainerVisual GetContainerVisualWith(double offsetY)
        {
            ContainerVisual contentVisual = new ContainerVisual();
            contentVisual.Transform = new TranslateTransform(0, offsetY);
            return contentVisual;
        }
    } 
}
