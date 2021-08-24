using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;

namespace ExcelTools
{
    [Table("tb_test")]
    public class tb_test
    {
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [MaxLength(36)]
        public string ID { get; set; }

        public string RID { get; set; }
        public string RBM { get; set; }
        public decimal SJ { get; set; }
        public int KC { get; set; }

        public bool JYBS { get; set; }

        public DateTime Addtime { get; set; }
        public string BZ { get; set; }
        public string MC { get; set; }
        public string PYM { get; set; }
        public string WBM { get; set; }
        public string PHONE { get; set; }
    }
}
