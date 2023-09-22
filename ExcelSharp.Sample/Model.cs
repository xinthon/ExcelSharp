using System.ComponentModel.DataAnnotations;

namespace ExcelSharp.Sample
{
    public class Model
    {
        [Display(Name = "EntityId")]
        [Width(Width = 10)]
        public string Id { get; set; } = string.Empty;

        [Display(Name = "EntityType")]
        [Width(Width = 20)]
        public string EntityType { get; set; } = string.Empty;

        public string CustomField1 { get; set; } = string.Empty;

        public string CustomField2 { get; set; } = string.Empty;

        public string CustomField3 { get; set; } = string.Empty;

        public string CustomField4 { get; set; } = string.Empty;

        public string CustomField5 { get; set; } = string.Empty;

        public string CustomField6 { get; set;} = string.Empty;

        public string CustomField7 { get; set; } = string.Empty;
    }
}
