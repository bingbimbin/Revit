using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using System.Windows;                       
using System.Windows.Media;                  
using System.Windows.Media.Imaging;

namespace NewJeans
{
    public class App : IExternalApplication
    {
        private const string TabName = "FAB_4";
 

        public Result OnStartup(UIControlledApplication uiApp)
        {
            try
            {
                // 탭 생성
                try
                {
                    uiApp.CreateRibbonTab(TabName);
                }
                catch (Exception) {}

                // 패널 생성
                RibbonPanel Connectorpanel = uiApp.CreateRibbonPanel(TabName, "Connectors");
                RibbonPanel Parameterpanel = uiApp.CreateRibbonPanel(TabName, "Parameters");
                RibbonPanel Familypanel = uiApp.CreateRibbonPanel(TabName, "Family Utils");


                // 버튼 아이콘 경로
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string DescriptionTextIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "DescriptionText.png");
                string ConnectorDuplicateIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "ConnectorDuplicate.png");
                string ExportConnectorDataIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "ExportConnectorData.png");
                string NozzleOrederIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "NozzleOrder.png");
                string DeleteFamilyParameterIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "DeleteFamilyParameter.png");
                string FamilyParameterIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "FamilyParameter.png");
                string LabelParameterIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "LabelParameter.png");
                string NameChangeIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "NameChange.png");
                string CreatePunchHoleIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "CreatePunchHole.png");
                string CreateConnectorsIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "CreateConnectors.png");
                string PlaceFamilyIconPath = Path.Combine(Path.GetDirectoryName(assemblyPath), "Resources", "PlaceFamily.png");

                // 4) 각 커맨드 등록

                PushButtonData btnDescriptionText = new PushButtonData(
                    "btnDescriptionText",
                    "Description\nText 배치",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "DescriptionTextNote.dll"),
                    "NewJeans.DescriptionText"
                )
                {
                    ToolTip = "세로로 패밀리 배치 후 가로로 타입 배치",
                    Image = GetScaledImage(DescriptionTextIconPath, 0.015625),
                    LargeImage = GetScaledImage(DescriptionTextIconPath, 0.03125)
                };

                PushButtonData btnConnectorDuplicate = new PushButtonData(
                    "btnConnectorDuplicate",            
                    "커넥터\n중복체크",         
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "ConnectorDuplicateCheck.dll"),
                    "NewJeans.ConnectorDuplicateCheck"
                )
                {
                    ToolTip = "패밀리를 먼저 프로젝트에 배치하세요.",
                    Image = GetScaledImage(ConnectorDuplicateIconPath, 0.25),
                    LargeImage = GetScaledImage(ConnectorDuplicateIconPath, 0.5)
                };


                PushButtonData btnExportConnectorData = new PushButtonData(
                    "btnExportConnectorData",
                    "커넥터 정보\n내보내기",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "ExportConnectorData.dll"),
                    "NewJeans.ExportConnectorData"
                )
                {
                    ToolTip = "검사대상 패밀리 파일 추가 후 진행하세요.",
                    Image = GetScaledImage(ExportConnectorDataIconPath, 0.25),
                    LargeImage = GetScaledImage(ExportConnectorDataIconPath, 0.5)
                };

                PushButtonData btnNozzleOreder = new PushButtonData(
                    "btnNozzleOreder",
                    "노즐 커넥터\n순서검토",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "NozzleOreder.dll"),
                    "NewJeans.NozzleOrderCheck"
                )
                {
                    ToolTip = "검사대상 패밀리를 프로젝트에 배치하세요.",
                    LongDescription = "평면상 왼쪽에서 오른쪽 순서로 커넥터 추출" +
                    "\nDescription 없으면 공란" +
                    "\n끝에 OO 붙으면 정상, !! 붙으면 확인필요, XX 붙으면 오류",
                    Image = GetScaledImage(NozzleOrederIconPath, 0.25),
                    LargeImage = GetScaledImage(NozzleOrederIconPath, 0.5)
                };

                PushButtonData btnDeleteFamilyParameter = new PushButtonData(
                    "btnDeleteFamilyParameter",
                    "패밀리 파라미터\n다중삭제",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "DeleteFamilyParameter.dll"),
                    "NewJeans.DeleteParameter"
                )
                {
                    ToolTip = "삭제 대상 패밀리를 프로젝트에 배치하세요",
                    LongDescription = "파라미터 입력은 다중으로 가능하고 쉼표로 구분해서 입력합니다." +
                    "\n화면 리스트에 나오는 패밀리들은 최상위 패밀리만",
                    Image = GetScaledImage(DeleteFamilyParameterIconPath, 0.25),
                    LargeImage = GetScaledImage(DeleteFamilyParameterIconPath, 0.5)
                };


                PushButtonData btnFamilyParameter = new PushButtonData(
                    "btnFamilyParameter",
                    "패밀리 파라미터\n내보내기",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "FamilyParameter.dll"),
                    "NewJeans.FamilyParameter"
                )
                {
                    ToolTip = "현재 프로젝트에 로드되어 있는 패밀리의 패밀리 파라미터 GUID와 함께 내보냅니다.",
                    Image = GetScaledImage(FamilyParameterIconPath, 0.25),
                    LargeImage = GetScaledImage(FamilyParameterIconPath, 0.5)
                };

                PushButtonData btnLabelParameter = new PushButtonData(
                    "btnLabelParameter",
                    "Label 파라미터\n내보내기",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "LableParameter.dll"),
                    "NewJeans.LabelParameter"
                )
                {
                    ToolTip = "현재 프로젝트에 로드되어 있는 패밀리의 Label 파라미터와 Family Type 파라미터 내보냅니다.",
                    Image = GetScaledImage(LabelParameterIconPath, 0.25),
                    LargeImage = GetScaledImage(LabelParameterIconPath, 0.5)
                };


                PushButtonData btnNameChange = new PushButtonData(
                    "btnNameChange",               // internal name
                    "패밀리&타입\n일괄변경",         // label
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "NameChange.dll"), // .dll 경로
                    "NewJeans.ChangeName"      // full class name of IExternalCommand
                )
                {
                    ToolTip = "Excel 기준으로 패밀리 및 타입명을 일괄 변경 후 저장합니다.",
                    LongDescription = 
                    "엑셀 해더는 [기존 패밀리명 / 기존 타입명 / 새 패밀리명 / 새 타입명] 이고 이 순서대로 데이터 입력" +
                    "\n4개의 열 다 입력하면 기존패밀리명과 타입명에 해당하는 걸 새 패밀리명과 타입명으로 수정" +
                    "\n[기존 타입명] 열을 비워놓으면 제일 첫번째 타입을 대상으로 수정" +
                    "\n[기존 타입명] 열에 *로 채워넣으면 타입명은 변경하시 않고 패밀리명만 변경" +
                    "\n엑셀 리스트에 있는데 추가안된 패밀리는 마지막에 메모장으로 나옵니다.",
                    Image = GetScaledImage(NameChangeIconPath, 0.25),
                    LargeImage = GetScaledImage(NameChangeIconPath, 0.5)
                };

                PushButtonData btnCreatePunchHole = new PushButtonData(
                    "btnCreatePunchHole",
                    "홀타공\nBase 생성",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "CreatePunchHole.dll"),
                    "NewJeans.CreatePunchHole"
                )
                {
                    ToolTip = "기준 Extrusion 상단에 맞춰 높이 150으로 생성",
                    LongDescription =
                    "홀타공 부분을 캐드에서 Layer 이름을 Punch_hole 로 설정하세요"+
                    "\n홀타공 도면을 평면에 배치하고 기준이 되는 Extrusion 선택하세요", 
                    Image = GetScaledImage(CreatePunchHoleIconPath, 0.25),
                    LargeImage = GetScaledImage(CreatePunchHoleIconPath, 0.5)
                };

                PushButtonData btnCreateConnectors = new PushButtonData(
                    "btnCreateConnectors",
                    "홀타공\nConnector 생성",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "CreateConnector.dll"),
                    "NewJeans.MakeConnectorsOnPlane"
                )
                {
                    ToolTip = "홀타공 면 위에 Connector 사이즈에 맞게 생성",
                    LongDescription =
                    "Connector 생성 후에 Extrusion을 void로 바꾸고 cut 해서 홀타공 생성",                     
                    Image = GetScaledImage(CreateConnectorsIconPath, 0.25),
                    LargeImage = GetScaledImage(CreateConnectorsIconPath, 0.5)
                };

                PushButtonData btnPlaceFamily = new PushButtonData(
                    "btnPlaceFamily",
                    "카테고리별\n패밀리 배치",
                    Path.Combine(Path.GetDirectoryName(assemblyPath), "PlaceFamily.dll"),
                    "NewJeans.PlaceFamily"
                )
                {
                    ToolTip = "세로로 패밀리 배치 후 가로로 타입 배치",
                    Image = GetScaledImage(PlaceFamilyIconPath, 0.015625),
                    LargeImage = GetScaledImage(PlaceFamilyIconPath, 0.03125)
                };

                Connectorpanel.AddItem(btnDescriptionText);
                Connectorpanel.AddItem(btnConnectorDuplicate);
                Connectorpanel.AddItem(btnExportConnectorData);
                Connectorpanel.AddItem(btnNozzleOreder);
                Parameterpanel.AddItem(btnDeleteFamilyParameter);
                Parameterpanel.AddItem(btnFamilyParameter);
                Parameterpanel.AddItem(btnLabelParameter);
                Familypanel.AddItem(btnNameChange);
                Familypanel.AddItem(btnCreatePunchHole);
                Familypanel.AddItem(btnCreateConnectors);
                Familypanel.AddItem(btnPlaceFamily);

                return Result.Succeeded;
            }


            catch (Exception ex)
            {
                TaskDialog.Show("OnStartup Error", ex.Message);
                return Result.Failed;
            }
        }

        private static ImageSource GetScaledImage(string path, double scale)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(path, UriKind.Absolute);
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();

            var scaled = new TransformedBitmap();
            scaled.BeginInit();
            scaled.Source = bitmap;
            scaled.Transform = new ScaleTransform(scale, scale); // ex: 0.25 = 64 → 16
            scaled.EndInit();

            return scaled;
        }


        public Result OnShutdown(UIControlledApplication uiApp)
        {
            return Result.Succeeded;
        }
    }
}
