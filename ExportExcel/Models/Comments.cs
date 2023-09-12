﻿using System.ComponentModel.DataAnnotations;

namespace ExportExcel.Models
{
    public class Comments
    {
        [Key]
        public int CommentID { get; set; }
        public string CommentUser { get; set; }
        public DateTime CommentDate { get; set; }
        public string CommentContent { get; set; }
        public bool CommentState { get; set; }
        public int DestinationId { get; set; }
    }
}
