using GalaSoft.MvvmLight;

namespace WeldCutList.ViewModel
{
    public class DrawingViewModel : ViewModelBase
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
    }
}
