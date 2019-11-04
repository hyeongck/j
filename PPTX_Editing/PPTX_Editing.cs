using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace PPTX_Class
{
    public class PPTX_Editing
    {
        public class YILED : INT
        {
            public PPT.Application pptApplication { get; set; }
            public PPT.Presentation pptPresentation { get; set; }
            public PPT.CustomLayout customLayout { get; set; }
            public PPT.Slides slides { get; set; }
            public PPT._Slide slide { get; set; }
            public PPT.TextRange objTex { get; set; }

            public void Open(string File)
            {
                pptApplication = new PPT.Application();
                pptPresentation = pptApplication.Presentations.Open(File, MsoTriState.msoFalse, MsoTriState.msoCTrue, MsoTriState.msoCTrue);

            }
            public void Slide(int Count)
            {
                slide = (PPT.Slide)pptApplication.ActivePresentation.Slides[1];
                slide.Copy();
                pptPresentation.Slides.Paste();

            }
            public void Title(string Text, string Name, int Size, int index)
            {
                objTex = slide.Shapes[1].TextFrame.TextRange;
                objTex.Text = Text;
              //  objTex.Font.Name = "Gulim";
                objTex.Font.Size = Size;



            }
            public void AddPicture(string File, float Left, float Top, float Width, float Height)
            {

                slide.Shapes.AddPicture("", MsoTriState.msoFalse, MsoTriState.msoTrue,
                150, 150, 500, 350);

            }
        }

        public class FCM_Automation_EXCEL : INT
        {
            public PPT.Application pptApplication { get; set; }
            public PPT.Presentation pptPresentation { get; set; }
            public PPT.CustomLayout customLayout { get; set; }
            public PPT.Slides slides { get; set; }
            public PPT._Slide slide { get; set; }
            public PPT.TextRange objTex { get; set; }

            public void Open(string File)
            {
                pptApplication = new PPT.Application();
                pptPresentation = pptApplication.Presentations.Open(File, MsoTriState.msoFalse, MsoTriState.msoCTrue, MsoTriState.msoCTrue);

            }
            public void Slide(int Count)
            {
                slide = (PPT.Slide)pptApplication.ActivePresentation.Slides[Count];

                slide.Copy();
                pptPresentation.Slides.Paste();

            }
            public void Title(string Text, string Name, int Size, int index)
            {

                PPT.CustomLayout customLayout =
                      pptPresentation.SlideMaster.CustomLayouts[PPT.PpSlideLayout.ppLayoutTextAndClipart];

                slide = pptPresentation.Slides.AddSlide(1, customLayout);


                objTex = slide.Shapes[1].TextFrame.TextRange;
                // objTex = pptPresentation.Slides[0].Shapes[1].TextFrame.TextRange;
      
                objTex.Text = Text;
                //  objTex.Font.Name = "Gulim";
                objTex.Font.Size = Size;
              

            }
            public void AddPicture(string File, float Left, float Top, float Width, float Height)
            {

                try
                {
                    slide.Shapes.AddPicture(File, MsoTriState.msoFalse, MsoTriState.msoTrue, Left, Top, Width, Height);
                }
                catch
                {

                }

            }
        }

        public interface INT
        {
            PPT.Application pptApplication { get; set; }
            PPT.Presentation pptPresentation { get; set; }
            PPT.CustomLayout customLayout { get; set; }
            PPT.Slides slides { get; set; }
            PPT._Slide slide { get; set; }
            PPT.TextRange objTex { get; set; }


            void Open(string File);
            void Slide(int Count);
            void Title(string Text, string Name, int Size, int index);
            void AddPicture(string File, float Left, float Top, float Width, float Height);

        }
        public INT Opened(string Key)
        {
            INT Int = null;
            switch (Key.ToUpper().Trim())
            {
                case "YIELD":
                    Int = new YILED();
                    break;
                case "FCM":
                    Int = new FCM_Automation_EXCEL();
                    break;

            }
            return Int;
        }
    }
}
