using EnvDTE;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace D365DeveloperExtensions.Core.Models
{
    public class ProjectListItem : INotifyPropertyChanged
    {
        private string _name;

        public Project Project { get; set; }
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