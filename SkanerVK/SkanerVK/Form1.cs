using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;
using System.Runtime.InteropServices;
using System.IO;

namespace SkanerVK
{


    public partial class Form1 : Form
    {

        static float scaleHeight = 11.70f;
        static float scaleWidth = 8.27f; // parametry skali dla obliczania pikseli w zaleznosci od dpi

        public Form1()
        {
            InitializeComponent();
        }           

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void ListScanners()
        {
            try
            {
                var deviceManager = new DeviceManager();

                //DeviceInfo 
                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                    {
                        continue;
                    }
                    else
                    {
                        //listBox1.Items.Add("Dupa");
                        listBox1.Items.Add(deviceManager.DeviceInfos[i].Properties["Name"].get_Value());
                        //break;
                    }
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ListScanners();
            textBox1.Text = "150";
            textBox2.Text = "0"; //kontrast
            bright.Text = "0";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                var deviceManager = new DeviceManager();
                DeviceInfo AvaiableScanner= null;
                int resolutionValue = int.Parse(textBox1.Text);
                int brightValue = int.Parse(bright.Text);
                int contrastValue = int.Parse(textBox2.Text);
                for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
                {
                    if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                    {
                        continue;
                    }
                    else
                    {
                        AvaiableScanner = deviceManager.DeviceInfos[i];
                        break; 
                    }
                }
               
                    var device = AvaiableScanner.Connect();
                    var scannerItem = device.Items[1];

                    int scanWidth = (int)(resolutionValue * scaleWidth); //jak nie dziala zamienic na int
                    int scanHeight = (int)(1.41 * scanWidth);
            
                    AdjustScannerSettings(scannerItem, resolutionValue, scanHeight,scanWidth, brightValue,contrastValue);
                    var imgfile = (ImageFile) scannerItem.Transfer(FormatID.wiaFormatJPEG);
                    var path = "C:\\Users\\lab\\Links\\KowolikKrol\\vkplik.jpeg";
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    imgfile.SaveFile(path);
                    pictureBox1.ImageLocation = path; //do modyfikacji
                }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private static void AdjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanHeight, int scanWidth, int bright, int contrast)
        {
            //const string WIA_SCAN_COLOR_MODE = "6146";
 
            const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
            const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
            const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
            const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
            const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
            const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
            const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
            const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
  
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, 0);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, 0);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeight);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidth);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidth);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, bright);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrast);

            
        }
        private static void SetWIAProperty(IProperties properties, object propertyName, object resolutionValue) //jak nie dziala to zamienic na object
        {
            try
            {
                Property property = properties.get_Item(ref propertyName);
                property.set_Value(ref resolutionValue);
            }
            catch (Exception ex) { };
        }
    }
}
