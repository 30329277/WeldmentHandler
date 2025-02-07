using System.ComponentModel;
using GalaSoft.MvvmLight;

namespace WeldCutList.ViewModel
{
    public class DrawingViewModel : ViewModelBase, INotifyPropertyChanged
    {
        private string sheetName;

        public string SheetName
        {
            get { return sheetName; }
            set { sheetName = value; RaisePropertyChanged(); }
        }

        private string viewName;

        public string ViewName
        {
            get { return viewName; }
            set { viewName = value; RaisePropertyChanged(); }
        }

        private bool _isMacro8Enabled;
        public bool IsMacro8Enabled
        {
            get => _isMacro8Enabled;
            set 
            {
                if (_isMacro8Enabled != value)
                {
                    _isMacro8Enabled = value;
                    RaisePropertyChanged(nameof(IsMacro8Enabled));
                }
            }
        }
    }
}
