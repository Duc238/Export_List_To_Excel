using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportListToExcel.Model
{
    public class ProfileActor:BaseViewModel
    {
        private int _Id;
        public int Id { get=> _Id; set { _Id = value;OnPropertyChanged(); } }
        private string _Name;
        public string Name { get => _Name; set { _Name = value; OnPropertyChanged(); } }
    }
}
