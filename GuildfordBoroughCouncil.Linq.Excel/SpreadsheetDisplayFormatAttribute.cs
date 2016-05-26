using System;

namespace GuildfordBoroughCouncil.Linq
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class SpreadsheetDisplayFormatAttribute : Attribute
    {
        //
        // Summary:
        //     Gets or sets a value that indicates whether empty string values ("") are automatically
        //     converted to null when the data field is updated in the data source.
        //
        // Returns:
        //     true if empty string values are automatically converted to null; otherwise, false.
        //     The default is true.
        public bool ConvertEmptyStringToNull { get; set; }
        //
        // Summary:
        //     Gets or sets the display format for the field value.
        //
        // Returns:
        //     A formatting string that specifies the display format for the value of the data
        //     field. The default is an empty string (""), which indicates that no special formatting
        //     is applied to the field value.
        public string DataFormatString { get; set; }
        //
        // Summary:
        //     Gets or sets the text that is displayed for a field when the field's value is
        //     null.
        //
        // Returns:
        //     The text that is displayed for a field when the field's value is null. The default
        //     is an empty string (""), which indicates that this property is not set.
        public string NullDisplayText { get; set; }

        public GemBox.Spreadsheet.HorizontalAlignmentStyle HorizontalAlignment { get; set; }

        public GemBox.Spreadsheet.VerticalAlignmentStyle VerticalAlignment { get; set; }
    }
}
