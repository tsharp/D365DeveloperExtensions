using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CrmDeveloperExtensions2.Core.Models
{
    public class MockingFrameworkListItem : INotifyPropertyChanged
    {
        private string _name;

        public MockingFramework MockingFramework { get; set; }
        public string Name
        {
            get => _name;
            set
            {
                _name = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}