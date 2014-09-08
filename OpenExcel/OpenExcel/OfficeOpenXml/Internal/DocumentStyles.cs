namespace OpenExcel.OfficeOpenXml.Internal
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    public class DocumentStyles
    {
        private WorkbookPart _wpart;
        private int b_count = -1;
        protected List<string> fillsXML = new List<string>();
        protected List<string> fontsXML = new List<string>();
        protected List<string> formatsXML = new List<string>();

        public DocumentStyles(WorkbookPart wpart)
        {
            this._wpart = wpart;
        }

        public Border borderCombine(Border elemNew, Border elemBase)
        {
            if (elemNew.TopBorder != null)
            {
                elemBase.TopBorder = (TopBorder) elemNew.TopBorder.CloneNode(true);
            }
            if (elemNew.BottomBorder != null)
            {
                elemBase.BottomBorder = (BottomBorder) elemNew.BottomBorder.CloneNode(true);
            }
            if (elemNew.LeftBorder != null)
            {
                elemBase.LeftBorder = (LeftBorder) elemNew.LeftBorder.CloneNode(true);
            }
            if (elemNew.RightBorder != null)
            {
                elemBase.RightBorder = (RightBorder) elemNew.RightBorder.CloneNode(true);
            }
            if (elemNew.DiagonalBorder != null)
            {
                elemBase.DiagonalBorder = (DiagonalBorder) elemNew.DiagonalBorder.CloneNode(true);
            }
            return elemBase;
        }

        protected bool compareBorder(Border bNew, Border bOld)
        {
            return this.GenericElementCompare(bNew, bOld);
        }

        protected bool compareBorderFake(Border bNew, Border bOld)
        {
            if (this.b_count <= 0)
            {
                this.b_count++;
                return false;
            }
            this.b_count = 0;
            return true;
        }

        protected bool compareFont(Font fNew, Font fBase)
        {
            return this.GenericElementCompare(fNew, fBase);
        }

        public uint EnsureCustomNumberingFormat(NumberingFormat nfNew)
        {
            Stylesheet stylesheet = this.EnsureStylesheet();
            if (stylesheet.NumberingFormats == null)
            {
                stylesheet.NumberingFormats = new NumberingFormats();
                stylesheet.Save();
            }
            if ((from nf in stylesheet.NumberingFormats.Elements<NumberingFormat>()
                where nf.FormatCode == nfNew.FormatCode
                select nf).FirstOrDefault<NumberingFormat>() == null)
            {
                uint num = (from nf in stylesheet.NumberingFormats.Elements<NumberingFormat>()
                    let id = nf.NumberFormatId ?? 0
                            select id).FirstOrDefault<UInt32Value>();
                uint num2 = Math.Max((uint) 0x80, (uint) (num + 1));
                nfNew.NumberFormatId = num2;
                stylesheet.NumberingFormats.Append(new OpenXmlElement[] { nfNew });
                stylesheet.NumberingFormats.Count = (UInt32Value)(uint)stylesheet.NumberingFormats.Count<OpenXmlElement>();
                stylesheet.Save();
            }
            return (uint) nfNew.NumberFormatId;
        }

        public Stylesheet EnsureStylesheet()
        {
            WorkbookPart part = this._wpart;
            if (part.WorkbookStylesPart == null)
            {
                part.AddNewPart<WorkbookStylesPart>();
                Stylesheet stylesheet = new Stylesheet();
                part.WorkbookStylesPart.Stylesheet = stylesheet;
                Font font2 = new Font();
                FontSize size = new FontSize {
                    Val = 11.0
                };
                font2.FontSize = size;
                Color color = new Color {
                    Theme = 1
                };
                font2.Color = color;
                FontName name = new FontName {
                    Val = "Calibri"
                };
                font2.FontName = name;
                FontFamilyNumbering numbering = new FontFamilyNumbering {
                    Val = 2
                };
                font2.FontFamilyNumbering = numbering;
                FontScheme scheme = new FontScheme {
                    Val = (FontSchemeValues)2
                };
                font2.FontScheme = scheme;
                Font font = font2;
                Fill fill3 = new Fill();
                PatternFill fill4 = new PatternFill {
                    PatternType = (PatternValues)0
                };
                fill3.PatternFill = fill4;
                Fill fill = fill3;
                Fill fill5 = new Fill();
                PatternFill fill6 = new PatternFill {
                    PatternType = (PatternValues)0x11
                };
                fill5.PatternFill = fill6;
                Fill fill2 = fill5;
                Border border = new Border {
                    LeftBorder = new LeftBorder(),
                    RightBorder = new RightBorder(),
                    TopBorder = new TopBorder(),
                    BottomBorder = new BottomBorder(),
                    DiagonalBorder = new DiagonalBorder()
                };
                CellFormat format = new CellFormat {
                    NumberFormatId = 0,
                    BorderId = 0,
                    FontId = 0,
                    FillId = 0
                };
                CellFormat format2 = new CellFormat {
                    NumberFormatId = 0,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    FormatId = 0
                };
                CellStyle style = new CellStyle {
                    Name = "Normal",
                    BuiltinId = 0,
                    FormatId = 0
                };
                Fonts fonts = new Fonts {
                    Count = 1
                };
                stylesheet.Fonts = fonts;
                stylesheet.Fonts.Append(new OpenXmlElement[] { font });
                Fills fills = new Fills {
                    Count = 2
                };
                stylesheet.Fills = fills;
                stylesheet.Fills.Append(new OpenXmlElement[] { fill });
                stylesheet.Fills.Append(new OpenXmlElement[] { fill2 });
                Borders borders = new Borders {
                    Count = 1
                };
                stylesheet.Borders = borders;
                stylesheet.Borders.Append(new OpenXmlElement[] { border });
                CellStyleFormats formats = new CellStyleFormats {
                    Count = 1
                };
                stylesheet.CellStyleFormats = formats;
                stylesheet.CellStyleFormats.Append(new OpenXmlElement[] { format });
                CellFormats formats2 = new CellFormats {
                    Count = 1
                };
                stylesheet.CellFormats = formats2;
                stylesheet.CellFormats.Append(new OpenXmlElement[] { format2 });
                CellStyles styles = new CellStyles {
                    Count = 1
                };
                stylesheet.CellStyles = styles;
                stylesheet.CellStyles.Append(new OpenXmlElement[] { style });
                stylesheet.DifferentialFormats = new DifferentialFormats();
                TableStyles styles2 = new TableStyles {
                    Count = 0
                };
                stylesheet.TableStyles = styles2;
            }
            return part.WorkbookStylesPart.Stylesheet;
        }

        protected Fill fillCombine(Fill elemNew, Fill elemBase)
        {
            if (elemNew.PatternFill != null)
            {
                elemBase.PatternFill = (PatternFill) elemNew.PatternFill.CloneNode(true);
                return elemBase;
            }
            if (elemNew.GradientFill != null)
            {
                elemBase.GradientFill = (GradientFill) elemNew.GradientFill.CloneNode(true);
            }
            return elemBase;
        }

        protected bool fillCompare(Fill fill1, Fill fill2)
        {
            bool flag = true;
            if (fill1.InnerXml != fill2.InnerXml)
            {
                flag = false;
            }
            return flag;
        }

        protected CellFormat formatCombine(CellFormat elemNew, CellFormat elemBase)
        {
            if ((elemNew.ApplyNumberFormat != null) && elemNew.ApplyNumberFormat.Value)
            {
                elemBase.NumberFormatId = elemNew.NumberFormatId;
                elemBase.ApplyNumberFormat = elemNew.ApplyNumberFormat;
            }
            if ((elemNew.ApplyFont != null) && elemNew.ApplyFont.Value)
            {
                elemBase.FontId = elemNew.FontId;
                elemBase.ApplyFont = elemNew.ApplyFont;
            }
            if ((elemNew.ApplyBorder != null) && elemNew.ApplyBorder.Value)
            {
                elemBase.BorderId = elemNew.BorderId;
                elemBase.ApplyBorder = elemNew.ApplyBorder;
            }
            if ((elemNew.ApplyFill != null) && elemNew.ApplyFill.Value)
            {
                elemBase.FillId = elemNew.FillId;
                elemBase.ApplyFill = elemNew.ApplyFill;
            }
            if (elemNew.FormatId != null)
            {
                elemBase.FormatId = elemNew.FormatId;
                return elemBase;
            }
            elemBase.FormatId = null;
            return elemBase;
        }

        protected bool formatCompare(CellFormat cfToTest, CellFormat cfExisting)
        {
            bool flag = true;
            if (((cfToTest.NumberFormatId != cfExisting.NumberFormatId) && ((cfToTest.NumberFormatId == null) || (cfExisting.NumberFormatId == null))) || (((cfToTest.NumberFormatId != null) && (cfExisting.NumberFormatId != null)) && (cfToTest.NumberFormatId.InnerText != cfExisting.NumberFormatId.InnerText)))
            {
                flag = false;
            }
            if (((cfToTest.FillId != cfExisting.FillId) && ((cfToTest.FillId == null) || (cfExisting.FillId == null))) || (((cfToTest.FillId != null) && (cfExisting.FillId != null)) && (cfToTest.FillId.InnerText != cfExisting.FillId.InnerText)))
            {
                flag = false;
            }
            if (((cfToTest.BorderId != cfExisting.BorderId) && ((cfToTest.BorderId == null) || (cfExisting.BorderId == null))) || (((cfToTest.BorderId != null) && (cfExisting.BorderId != null)) && (cfToTest.BorderId.InnerText != cfExisting.BorderId.InnerText)))
            {
                flag = false;
            }
            if (((cfToTest.FontId != cfExisting.FontId) && ((cfToTest.FontId == null) || (cfExisting.FontId == null))) || (((cfToTest.FontId != null) && (cfExisting.FontId != null)) && (cfToTest.FontId.InnerText != cfExisting.FontId.InnerText)))
            {
                flag = false;
            }
            return ((((cfToTest.FormatId == cfExisting.FormatId) || ((cfToTest.FormatId != null) && (cfExisting.FormatId != null))) && (((cfToTest.FormatId == null) || (cfExisting.FormatId == null)) || !(cfToTest.FormatId.InnerText != cfExisting.FormatId.InnerText))) && flag);
        }

        private bool GenericElementCompare(OpenXmlElement e1, OpenXmlElement e2)
        {
            return e1.InnerXml.Equals(e2.InnerXml);
        }

        public Border GetBorder(uint idx)
        {
            return this._wpart.WorkbookStylesPart.Stylesheet.Borders.Elements<Border>().ElementAt<Border>(((int) idx));
        }

        public CellFormat GetCellFormat(uint idx)
        {
            return this._wpart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ElementAt<CellFormat>(((int) idx));
        }

        public Fill GetFill(uint idx)
        {
            return this._wpart.WorkbookStylesPart.Stylesheet.Fills.Elements<Fill>().ElementAt<Fill>(((int) idx));
        }

        protected string getFillHash(Fill fill)
        {
            return fill.InnerXml;
        }

        public Font GetFont(uint idx)
        {
            Stylesheet stylesheet = this._wpart.WorkbookStylesPart.Stylesheet;
            if ((stylesheet != null) && (stylesheet.Fonts != null))
            {
                return stylesheet.Fonts.Elements<Font>().ElementAt<Font>(((int) idx));
            }
            return null;
        }

        protected string getFontHash(Font fnt)
        {
            return fnt.InnerXml;
        }

        protected string getFormatHash(CellFormat format)
        {
            return ((format.NumberFormatId) + "|" + (format.FillId) + "|" + (format.BorderId) + "|" + (format.FontId) + "|" + (format.FormatId));
        }

        public NumberingFormat GetNumberingFormat(uint numFmtId)
        {
            Stylesheet stylesheet = this._wpart.WorkbookStylesPart.Stylesheet;
            if (stylesheet.NumberingFormats == null)
            {
                stylesheet.NumberingFormats = new NumberingFormats();
                stylesheet.Save();
            }
            return (from x in stylesheet.NumberingFormats.Elements<NumberingFormat>()
                where x.NumberFormatId == numFmtId
                select x).FirstOrDefault<NumberingFormat>();
        }

        public bool IsDateFormat(CellFormat cf)
        {
            return ((cf.NumberFormatId >= 14) && (cf.NumberFormatId <= 0x16));
        }

        public uint MergeAndRegisterBorder(Border bNew, UInt32Value baseBordersIdx, bool doSave)
        {
            uint num;
            Stylesheet stylesheet = this.EnsureStylesheet();
            if (baseBordersIdx == "0")
            {
                num = this.MergeAndRegisterStyleElement<Border, Borders>(bNew, stylesheet.Borders, new Func<Border, Border, Border>(this.borderCombine), new Func<Border, Border, bool>(this.compareBorderFake), baseBordersIdx, doSave);
            }
            else
            {
                num = this.MergeAndRegisterStyleElement<Border, Borders>(bNew, stylesheet.Borders, new Func<Border, Border, Border>(this.borderCombine), new Func<Border, Border, bool>(this.compareBorder), baseBordersIdx, doSave);
            }
            if (stylesheet.Borders.Count != stylesheet.Borders.Count<OpenXmlElement>())
            {
                stylesheet.Borders.Count = (UInt32Value) (uint)stylesheet.Borders.Count<OpenXmlElement>();
                if (doSave)
                {
                    stylesheet.Save();
                }
            }
            return num;
        }

        public uint MergeAndRegisterCellFormat(CellFormat cfNew, UInt32Value baseCellXfsIdx, bool doSave)
        {
            if (cfNew.NumberFormatId == null)
            {
                cfNew.NumberFormatId = 0;
            }
            Stylesheet stylesheet = this.EnsureStylesheet();
            uint num = this.MergeAndRegisterStyleElement<CellFormat, CellFormats>(cfNew, stylesheet.CellFormats, new Func<CellFormat, CellFormat, CellFormat>(this.formatCombine), new Func<CellFormat, CellFormat, bool>(this.formatCompare), baseCellXfsIdx, doSave);
            if (stylesheet.CellFormats.Count != stylesheet.CellFormats.Count<OpenXmlElement>())
            {
                stylesheet.CellFormats.Count = (UInt32Value)(uint)stylesheet.CellFormats.Count<OpenXmlElement>();
                if (doSave)
                {
                    stylesheet.Save();
                }
            }
            return num;
        }

        public uint MergeAndRegisterFill(Fill fNew, UInt32Value baseFillsIdx, bool doSave)
        {
            Stylesheet stylesheet = this.EnsureStylesheet();
            uint num = this.MergeAndRegisterStyleElement<Fill, Fills>(fNew, stylesheet.Fills, new Func<Fill, Fill, Fill>(this.fillCombine), new Func<Fill, Fill, bool>(this.fillCompare), baseFillsIdx, doSave);
            if (stylesheet.Fills.Count != stylesheet.Fills.Count<OpenXmlElement>())
            {
                stylesheet.Fills.Count = (UInt32Value)(uint)stylesheet.Fills.Count<OpenXmlElement>();
                if (doSave)
                {
                    stylesheet.Save();
                }
            }
            return num;
        }

        public uint MergeAndRegisterFont(Font fNew, UInt32Value baseFontsIdx, bool doSave)
        {
            Stylesheet stylesheet = this.EnsureStylesheet();
            uint num = this.MergeAndRegisterStyleElement<Font, Fonts>(fNew, stylesheet.Fonts, new Func<Font, Font, Font>(this.MergeFont), new Func<Font, Font, bool>(this.compareFont), baseFontsIdx, doSave);
            if (stylesheet.Fonts.Count != stylesheet.Fonts.Count<OpenXmlElement>())
            {
                stylesheet.Fonts.Count = (UInt32Value)(uint)stylesheet.Fonts.Count<OpenXmlElement>();
                if (doSave)
                {
                    stylesheet.Save();
                }
            }
            return num;
        }

        private uint MergeAndRegisterStyleElement<TElement, TParent>(TElement elemNew, TParent parent, Func<TElement, TElement, TElement> fnCombine, Func<TElement, TElement, bool> fnCompare, UInt32Value baseElementIdx, bool doSave) where TElement: OpenXmlElement where TParent: OpenXmlCompositeElement
        {
            int num = -1;
            TElement local = default(TElement);
            if (baseElementIdx != null)
            {
                local = (TElement) parent.Elements<TElement>().ElementAt<TElement>((int)baseElementIdx.Value).Clone();
                local = fnCombine(elemNew, local);
            }
            else
            {
                local = elemNew;
            }
            int num2 = 0;
            bool flag = local is Font;
            bool flag2 = local is CellFormat;
            bool flag3 = local is Fill;
            string item = "";
            if ((flag || flag2) || flag3)
            {
                if (flag)
                {
                    item = this.getFontHash(local as Font);
                }
                else if (flag2)
                {
                    item = this.getFormatHash(local as CellFormat);
                }
                else
                {
                    item = this.getFillHash(local as Fill);
                }
            }
            List<string> fontsXML = new List<string>();
            if (flag)
            {
                fontsXML = this.fontsXML;
            }
            else if (flag2)
            {
                fontsXML = this.formatsXML;
            }
            else
            {
                fontsXML = this.fillsXML;
            }
            foreach (TElement local2 in parent.Elements<TElement>())
            {
                if (!flag && !flag2)
                {
                    if (!fnCompare(local2, local))
                    {
                        goto Label_01B4;
                    }
                    num = num2;
                    break;
                }
                if (fontsXML.Count <= num2)
                {
                    if (flag)
                    {
                        fontsXML.Add(this.getFontHash(local2 as Font));
                    }
                    else if (flag2)
                    {
                        fontsXML.Add(this.getFormatHash(local2 as CellFormat));
                    }
                    else
                    {
                        fontsXML.Add(this.getFillHash(local2 as Fill));
                    }
                }
                if (item.Equals(fontsXML[num2]))
                {
                    num = num2;
                    break;
                }
            Label_01B4:
                num2++;
            }
            if (num == -1)
            {
                if ((flag || flag2) || flag3)
                {
                    fontsXML.Add(item);
                }
                parent.Append(new OpenXmlElement[] { local });
                if (doSave)
                {
                    this.EnsureStylesheet().Save();
                }
                num = parent.ChildElements.Count - 1;
            }
            return (uint) num;
        }

        public Font MergeFont(Font fontNew, Font fontTarget)
        {
            if (fontNew.FontCharSet != null)
            {
                fontTarget.FontCharSet.Val = fontNew.FontCharSet.Val;
            }
            if (fontNew.FontFamilyNumbering != null)
            {
                fontTarget.FontFamilyNumbering.Val = fontNew.FontFamilyNumbering.Val;
            }
            if (fontNew.FontName != null)
            {
                fontTarget.FontName.Val = fontNew.FontName.Val;
            }
            if (fontNew.FontScheme != null)
            {
                fontTarget.FontScheme.Val = fontNew.FontScheme.Val;
            }
            if (fontNew.FontSize != null)
            {
                fontTarget.FontSize.Val = fontNew.FontSize.Val;
            }
            if (fontNew.Bold != null)
            {
                if (fontNew.Bold.Val == "0")
                {
                    fontTarget.Bold = null;
                }
                else
                {
                    fontTarget.Bold = fontTarget.Bold ?? new Bold();
                    fontTarget.Bold.Val = fontNew.Bold.Val;
                }
            }
            else
            {
                fontTarget.Bold = null;
            }
            if (fontNew.Italic != null)
            {
                if (fontNew.Italic.Val == "0")
                {
                    fontTarget.Italic = null;
                    return fontTarget;
                }
                fontTarget.Italic = fontTarget.Italic ?? new Italic();
                fontTarget.Italic.Val = fontNew.Italic.Val;
            }
            return fontTarget;
        }

        public void Save()
        {
            this._wpart.WorkbookStylesPart.Stylesheet.Save();
        }
    }
}

