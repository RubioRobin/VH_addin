using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;

namespace VH_Addin.Models
{
    public class ParameterInput : INotifyPropertyChanged
    {
        private string _value;

        public string Name { get; set; }
        
        public ObservableCollection<string> Options { get; set; } = new ObservableCollection<string>();

        public string Value 
        { 
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
