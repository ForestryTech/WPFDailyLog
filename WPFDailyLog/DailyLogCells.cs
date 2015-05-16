using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFDailyLog
{
    class DailyLogCells
    {   
        /* This class is the information for the cell that is getting created for the worksheet.
         * range is the range that will get merged and it is MANDATORY, and the cell will hold the text or formula
         * cell will be locked unless no cell contents are entered, then by default the cell will be unlocked
         * if no contents and cell needs to be locked, then there is a construtor to do this.
         */
        private string range;
        private string cellContents;
        private bool isFormula;
        private CellText cellText;
        private bool cellLocked;
        private CellBorder cellBorder;

        public string Range { get { return range; } }
        public string CellContents { get { return cellContents; }  }
        public bool IsFormula { get { return isFormula; } }
        public CellText CellText { get { return cellText; } }
        public bool CellLocked { get { return cellLocked; } }
        public CellBorder CellBorder { get { return cellBorder; } }
        public bool Hidden { // If cell contains a formula hide it
            get {
                if (isFormula)
                    return true;
                else
                    return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellContents"></param>
        /// <param name="isFormula"></param>
        /// <param name="cellText"></param>
        public DailyLogCells(string range, string cellContents, bool isFormula, CellText cellText) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = cellText;
            this.isFormula = isFormula;
            this.cellLocked = true;
            this.cellBorder = CellBorder.Yes;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellContents"></param>
        /// <param name="isFormula"></param>
        public DailyLogCells(string range, string cellContents, bool isFormula) {
            // Constructor mainly used for formulas as no text formating is needed.
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = new CellText();
            this.isFormula = isFormula;
            this.cellLocked = true;
            this.cellBorder = CellBorder.Yes;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellContents"></param>
        public DailyLogCells(string range, string cellContents) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = new CellText();
            this.isFormula = false;
            this.cellLocked = true;
            this.cellBorder = CellBorder.Yes;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        public DailyLogCells(string range) {
            // mainly used for empty cells. Ones that get filled in when user is entering in daily log
            this.range = range;
            this.cellContents = "";
            this.cellText = new CellText();
            this.isFormula = false;
            this.cellLocked = false;
            this.cellBorder = CellBorder.Yes;
        }
        public DailyLogCells(string range, bool cellLocked) {
            this.range = range;
            this.cellContents = "";
            this.cellText = new CellText();
            this.isFormula = false;
            this.cellLocked = cellLocked;
            this.cellBorder = CellBorder.Yes;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellContents"></param>
        /// <param name="cellText"></param>
        public DailyLogCells(string range, string cellContents, CellText cellText) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = cellText;
            this.isFormula = false;
            this.cellLocked = true;
            this.cellBorder = CellBorder.Yes;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellText"></param>
        public DailyLogCells(string range, CellText cellText) {
            this.range = range;
            this.cellContents = "";
            this.cellText = cellText;
            this.isFormula = false;
            this.cellLocked = false;
            this.cellBorder = CellBorder.Yes;
        }

        public DailyLogCells(string range, CellText cellText, bool cellLocked) {
            this.range = range;
            this.cellContents = "";
            this.cellText = cellText;
            this.isFormula = false;
            this.cellLocked = cellLocked;
            this.cellBorder = CellBorder.Yes;
        }

        public DailyLogCells(string range, string cellContents, bool isFormula, CellText cellText, CellBorder cellBorder) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = cellText;
            this.isFormula = isFormula;
            this.cellLocked = true;
            this.cellBorder = cellBorder;
        }

        public DailyLogCells(string range, string cellContents, CellText cellText, CellBorder cellBorder) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = cellText;
            this.isFormula = false;
            this.cellLocked = true;
            this.cellBorder = cellBorder;
        }

        public DailyLogCells(string range, string cellContents, bool isFormula, CellBorder cellBorder) {
            this.range = range;
            this.cellContents = cellContents;
            this.cellText = new CellText();
            this.isFormula = isFormula;
            this.cellLocked = false;
            this.cellBorder = cellBorder;
        }

    }

    class CellText
    {
        public static int DEFAULT_FONT_SIZE = 11;
        public FontStyle Style;
        public FontPosition Position;
        public FontColor Color;
        public int Size;

        public CellText(FontStyle style, FontPosition position, FontColor color, int size) {
            this.Style = style;
            this.Position = position;
            this.Color = color;
            this.Size = size;
        }

        // 1 parameers in constructor
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        public CellText(FontStyle style) : this(style, FontPosition.Center, FontColor.Black, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="position"></param>
        public CellText(FontPosition position) : this(FontStyle.Regular, position, FontColor.Black, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="color"></param>
        public CellText(FontColor color) : this(FontStyle.Regular, FontPosition.Center, color, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="size"></param>
        public CellText(int size) : this(FontStyle.Regular, FontPosition.Center, FontColor.Black, size) {}
        // 3 parameters in constructor
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="position"></param>
        /// <param name="color"></param>
        public CellText(FontStyle style, FontPosition position, FontColor color) : this(style, position, color, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="size"></param>
        public CellText(FontStyle style, FontColor color, int size) : this(style, FontPosition.Center, color, size) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="position"></param>
        /// <param name="size"></param>
        public CellText(FontStyle style, FontPosition position, int size) : this(style, position, FontColor.Black, size) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="position"></param>
        /// <param name="color"></param>
        /// <param name="size"></param>
        public CellText(FontPosition position, FontColor color, int size) : this(FontStyle.Regular, position, color, size) { }
        // 2 parameters in constructor
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="position"></param>
        public CellText(FontStyle style, FontPosition position) : this(style, position, FontColor.Black, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public CellText(FontStyle style, FontColor color) : this(style, FontPosition.Center, color, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="size"></param>
        public CellText(FontStyle style, int size) : this(style, FontPosition.Center, FontColor.Black, size) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="position"></param>
        /// <param name="color"></param>
        public CellText(FontPosition position, FontColor color) : this(FontStyle.Regular, position, color, DEFAULT_FONT_SIZE) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="position"></param>
        /// <param name="size"></param>
        public CellText(FontPosition position, int size) : this(FontStyle.Regular, position, FontColor.Black, size) { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="color"></param>
        /// <param name="size"></param>
        public CellText(FontColor color, int size) : this(FontStyle.Regular, FontPosition.Center, color, size) { }
        // empty constructor
        /// <summary>
        /// All default value for CellText()
        /// </summary>
        public CellText() : this(FontStyle.Regular, FontPosition.Center, FontColor.Black, DEFAULT_FONT_SIZE) { }
    }
    enum FontStyle
    {
        Regular,
        Bold,
        Italic,
        BoldItalic,
    }

    enum FontPosition
    {
        Left,
        Center,
        Right,
    }

    enum FontColor
    {
        Red,
        Black,
        Green,
    }

    enum CellBorder
    {
        Yes,
        No,
    }
}
