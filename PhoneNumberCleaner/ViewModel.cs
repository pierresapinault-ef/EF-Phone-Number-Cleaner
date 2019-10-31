using System.ComponentModel;

namespace PhoneNumberCleaner
{
    public class ViewModel : INotifyPropertyChanged
    {
        public ViewModel()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        internal void NotifyPropertyChanged(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }
}