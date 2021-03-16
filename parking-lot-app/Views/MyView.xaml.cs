using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Input;

namespace parking_lot_app.Views
{
    /// <summary>
    /// Interaction logic for MyView
    /// </summary>
    public partial class MyView : UserControl
    {
        public MyView()
        {
            InitializeComponent();
            TextCompositionManager.AddPreviewTextInputStartHandler(this.A, new TextCompositionEventHandler(NumberValidationTextBox));
            TextCompositionManager.AddPreviewTextInputStartHandler(this.B, new TextCompositionEventHandler(NumberValidationTextBox));
            TextCompositionManager.AddPreviewTextInputStartHandler(this.C, new TextCompositionEventHandler(NumberValidationTextBox));
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }
    }
}