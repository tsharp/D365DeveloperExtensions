using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace WebResourceDeployer.ViewModels
{
    public class WebResourceItem : INotifyPropertyChanged
    {
        private bool _publish;
        public bool Publish
        {
            get => _publish;
            set
            {
                if (_publish == value) return;

                _publish = value;
                OnPropertyChanged();
            }
        }
        public Guid WebResourceId { get; set; }
        public int Type { get; set; }
        public string TypeName { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public bool IsManaged { get; set; }
        private ObservableCollection<MenuItem> _projectFolders;
        public ObservableCollection<MenuItem> ProjectFolders
        {
            get => _projectFolders;
            set
            {
                if (_projectFolders == value) return;

                _projectFolders = value;
                OnPropertyChanged();
            }
        }
        public bool AllowCompare => SetAllowCompare();
        public bool AllowPublish => SetAllowPublish();
        private string _boundFile;
        public string BoundFile
        {
            get => _boundFile;
            set
            {
                if (_boundFile == value) return;

                _boundFile = value;
                OnPropertyChanged();
                OnPropertyChanged("AllowPublish");
            }
        }
        public Guid SolutionId { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private bool SetAllowCompare()
        {
            if (string.IsNullOrEmpty(BoundFile))
                return false;

            int[] noCompare = { 5, 6, 7, 8, 10 };
            return !noCompare.Contains(Type);
        }

        private bool SetAllowPublish()
        {
            if (IsManaged)
            {
                Publish = false;
                return false;
            }
            if (string.IsNullOrEmpty(BoundFile))
            {
                Publish = false;
                return false;
            }
            return true;
        }
    }
}
