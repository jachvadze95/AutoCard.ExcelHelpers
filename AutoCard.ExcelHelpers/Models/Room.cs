using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCard.ExcelHelpers.Models
{
    public class Room
    {
        public int RoomNo { get; set; }
        public string RoomName { get; set; }
        public double Height { get; set; }

        public List<RoomWall> Walls { get; set; } = new List<RoomWall>();
    }
}
