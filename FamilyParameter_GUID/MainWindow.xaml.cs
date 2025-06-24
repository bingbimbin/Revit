using Autodesk.Revit.UI;
using System.Windows;
using System.Windows.Controls;

namespace NewJeans
{
    public partial class MainUserControl : UserControl
    {
        public MainUserControl()
        {
            InitializeComponent();

            // 핸들러 및 이벤트 초기화
            var handler1 = new FamilyParameterHandler();
            var event1 = ExternalEvent.Create(handler1);

            var handler2 = new FamilyParameterHandler2();
            var event2 = ExternalEvent.Create(handler2);

            // 컨트롤 생성 및 초기화
            var control1 = new FamilyParameterResult();
            control1.Initialize(event1, handler1);

            var control2 = new FamilyParameterResult2();
            control2.Initialize(event2, handler2);

            // 각 TabItem에 컨트롤 삽입
            tabProjectFamily.Content = control1;
            tabRfaFamily.Content = control2;
        }
    }
}
