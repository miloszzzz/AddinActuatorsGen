using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddInActuatorsGen.Models
{
    public class Comment
    {
        public string CommentText { get; set; }

        public Comment(string commentText)
        {
            CommentText = commentText;
        }
    }
}
