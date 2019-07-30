using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using TemplateWizards.Models;

namespace TemplateWizards
{
    public partial class CustomTemplatePicker
    {
        private ObservableCollection<CustomTemplateListItem> _templates;

        public ObservableCollection<CustomTemplateListItem> Templates
        {
            get => _templates;
            set
            {
                if (value != null && _templates == value)
                    return;

                _templates = value;
                OnPropertyChanged();
            }
        }
        public CustomTemplate SelectedTemplate;

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public CustomTemplatePicker(List<CustomTemplate> templates)
        {
            InitializeComponent();
            DataContext = this;

            Templates = new ObservableCollection<CustomTemplateListItem>();

            DisplayTemplates(templates);
        }

        private void DisplayTemplates(List<CustomTemplate> templates)
        {
            templates.ForEach(delegate (CustomTemplate template)
            {
                Templates.Add(CreateCustomTemplateListItem(template));
            });
        }

        private static CustomTemplateListItem CreateCustomTemplateListItem(CustomTemplate template)
        {
            return new CustomTemplateListItem
            {
                Name = template.DisplayName,
                Template = template,
                Description = template.Description,
                Selected = false
            };
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CloseDialog(false);
        }

        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            CloseDialog(true);
        }

        private void CloseDialog(bool result)
        {
            DialogResult = result;
            Close();
        }

        private void LanguageTemplates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count < 1)
            {
                SelectedTemplate = null;
                Ok.IsEnabled = false;
                return;
            }

            var listItem = (CustomTemplateListItem)e.AddedItems[0];
            SelectedTemplate = listItem.Template;
            Ok.IsEnabled = true;
        }
    }
}