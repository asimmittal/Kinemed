using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using WindowsInput;
using System.Threading;
using Microsoft.Research.Kinect.Nui;

namespace WpfApplication1
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Point3D class basically defines a point as a collection of cartesian coordinates and depth values. This
        /// will be used to represent points extracted from the depth map
        /// </summary>
        class Point3D
        {
            public double x, y;
            public int depth;
            public bool isValid;

            public Point3D()
            {
                isValid = false;
            }
        };

        //NUI objects
        Runtime nui;
        Point3D ptHandRight, ptHandLeft,ptHead;
        Point3D ptReference;
        bool isHandUp = false;

        //various constants that are needed
        const int WIDTH = 320;                                      //width of the depth map
        const int HEIGHT = 240;                                     //height of the depth map
        const int MAX_COLOR = 255;                                  //maximum value for color of a single Rgb channel
        const int MAX_ZONES = 20;                                   //maximum number of color zones 
        const int DEPTH_THRESH_ONE = 3000;                          //difference in depth (between head/spine and hand) that is needed to triger selection of either forward / backward
        const int DEPTH_THRESH_TWO = 6000;                          //minimum difference in depth needed between each hand and the head to trigger actions for both hands
        const int BPP = 3;                                          //number of bytes that will be used per pixel to draw the depth map        

        //this imports the keyboard event routines. will be used to generate key down events - pulled this code off social.msdn.com
        [DllImport("user32.dll")]
        public static extern void keybd_event(
        byte bVk,
        byte bScan,
        uint dwFlags,
        uint dwExtraInfo
        );

        //key codes for left arrow, down arrow and keyUP trigger
        const int VK_LEFT = 0x25;
        const int VK_RIGHT = 0x27;
        const uint KEYEVENTF_KEYUP = 0x2;

        //simple structure to define a cartesian point
        public struct PointAPI
        {
            public int X;
            public int Y;
        }

        //win32 calls to set and get cursor positions
        public partial class NativeMethods
        {
            /// Return Type: BOOL->int  
            ///X: int  
            ///Y: int  
            [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "SetCursorPos")]
            [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public static extern bool SetCursorPos(int X, int Y);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetCursorPos(out PointAPI lpPoint);                       
        }

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010
        }

        
        void eventLeftArrowDown()   {keybd_event((byte)VK_LEFT, 0, 0, 0);}
        void eventLeftArrowUp()     {keybd_event((byte)VK_LEFT, 0, KEYEVENTF_KEYUP, 0);}
        void eventRightArrowDown()  {keybd_event((byte)VK_RIGHT, 0, 0, 0);}
        void eventRightArrowUp()    {keybd_event((byte)VK_RIGHT, 0, KEYEVENTF_KEYUP, 0);}

        bool isLeftMouseDown = false;
        void eventMouseLeftDown()   
        {
            if (!isLeftMouseDown)
            {
                mouse_event((int)(MouseEventFlags.LEFTDOWN), 0, 0, 0, 0);
                isLeftMouseDown = true;
            }
        }

        void eventMouseLeftUp()
        {
            mouse_event((int)(MouseEventFlags.LEFTUP), 0, 0, 0, 0);
            isLeftMouseDown = false;
        }

        void eventMoveMouseToCenterScreen()
        {
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            NativeMethods.SetCursorPos((int)(screenWidth / 2), (int)(screenHeight / 2));
        }
        
        /// <summary>
        /// The class contructor - creates the UI 
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();                                  //create the UI for the window
                                   
        }

        /// <summary>
        /// This is called when the window is loaded. We use the runtime object configure the data format we need from the sensor
        /// and initialize it. We also bind event handlers to process depth and skeletal frames
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            depthMap.Width = WIDTH;                                 //set the width of the depth map image control
            depthMap.Height = HEIGHT;                               //set the heigh of the depth map image control

            nui = new Runtime();                                    //create a nui runtime
            ptHead = new Point3D();                                 //intialize the point that will track the head/spine
            ptHandRight = new Point3D();                            //initialize the point that will track the hand
            ptHandLeft = new Point3D();
            ptReference = new Point3D();
            nullfiyReference();

            this.Top = 0;
            this.Left = SystemParameters.PrimaryScreenWidth - this.Width;
            
            try
            {
                nui.Initialize(RuntimeOptions.UseDepthAndPlayerIndex | RuntimeOptions.UseSkeletalTracking | RuntimeOptions.UseColor);
                nui.VideoStream.Open(ImageStreamType.Video, 2, ImageResolution.Resolution640x480, ImageType.Color);
                nui.DepthStream.Open(ImageStreamType.Depth, 2, ImageResolution.Resolution320x240, ImageType.DepthAndPlayerIndex);
                nui.NuiCamera.ElevationAngle = 10;                
            }
            catch (InvalidOperationException)
            {
                System.Windows.MessageBox.Show("Runtime initialization failed. Please make sure Kinect device is plugged in and the image type being requested is supported.");
                this.Close();
            }

            nui.DepthFrameReady += new EventHandler<ImageFrameReadyEventArgs>(nui_DepthFrameReady);
            nui.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(nui_SkeletonFrameReady);
            nui.VideoFrameReady += new EventHandler<ImageFrameReadyEventArgs>(nui_VideoFrameReady);
        }
              

        /// <summary>
        /// the logic for action to be taken when no hand is up is as follows
        /// </summary>
        void performActionForNoHands()
        {
            LEDLeft.Fill = Brushes.Transparent;
            LEDRight.Fill = Brushes.Transparent;
            eventLeftArrowUp();
            eventRightArrowUp();
            chkMouseLock.IsChecked = false;
        }

        /// <summary>
        /// The logic for the action to be taken when the right hand is stretched out goes in here
        /// </summary>
        /// <param name="handPoint"></param>
        /// <param name="headPoint"></param>
        void performActionForRightHand(Point3D handPoint, Point3D headPoint)
        {
            if (handPoint.x < headPoint.x)
            {
                //activate left marker
                LEDLeft.Fill = Brushes.Green;
                LEDRight.Fill = Brushes.Transparent;
                eventLeftArrowDown();
            }
            else
            {
                //activate right marker
                LEDRight.Fill = Brushes.Green;
                LEDLeft.Fill = Brushes.Transparent;
                eventRightArrowDown();
            }
        }

        /// <summary>
        /// The logic for the action to be taken when the left hand is stretched out goes in here
        /// </summary>
        /// <param name="handPoint"></param>
        /// <param name="headPoint"></param>
        void performActionForLeftHand(Point3D handPoint, Point3D headPoint)
        {
            if (handPoint.x < headPoint.x)
            {
                //activate left marker
                LEDLeft.Fill = Brushes.Green;
                LEDRight.Fill = Brushes.Transparent;
                //eventLeftArrowDown();
                //eventLeftArrowUp();
                //Thread.Sleep(20);
            }
            else
            {
                //activate right marker
                LEDRight.Fill = Brushes.Green;
                LEDLeft.Fill = Brushes.Transparent;
                //eventRightArrowDown();
                //eventRightArrowUp();
                //Thread.Sleep(20);
            }
        }

        /// <summary>
        /// logic for the action to be performed when both hands are raised goes in here
        /// </summary>
        /// <param name="leftHand"></param>
        /// <param name="rightHand"></param>
        /// <param name="head"></param>
        void performActionForTwoHands(Point3D leftHand, Point3D rightHand, Point3D head)
        {
            LEDLeft.Fill = Brushes.OrangeRed;
            LEDRight.Fill = Brushes.OrangeRed;
            int scale = 8;

            Point3D center = new Point3D();
            center.x = (leftHand.x + rightHand.x) / 2;
            center.y = (leftHand.y + rightHand.y) / 2;

            eventMouseLeftDown();
            chkMouseLock.IsChecked = true;

            if (ptReference.x != -1)
            {
                int xDel = (int)(center.x - ptReference.x);
                int yDel = (int)(center.y - ptReference.y);

                PointAPI pt = new PointAPI();
                NativeMethods.GetCursorPos(out pt);
                pt.X += xDel*scale; pt.Y += yDel*scale;
                NativeMethods.SetCursorPos((int)pt.X, (int)pt.Y);
            }
            
            ptReference.x = center.x;
            ptReference.y = center.y;
            
        }



        /// <summary>
        /// 
        /// </summary>
        void nullfiyReference()
        {
            ptReference.x = -1; ptReference.y = -1;
            eventMouseLeftUp();

            if (chkMouseLock.IsChecked == true)
            {
                eventMoveMouseToCenterScreen();
            }
        }

        /// <summary>
        /// This routine is where the difference in depth between hand and head (or spine) is calculated
        /// based on this difference we decide whether the user wants to control the system or not. If he does 
        /// want to control, we check the position of his hand (either left of center or right of center).
        /// 
        /// if the hand is to the right of center, we activate the right arrow down event else the left arrow down event.
        /// Accordingly the required ui control on the screen is also highlighted
        /// </summary>
        /// <param name="head">point that tracks the head/spine</param>
        /// <param name="hand">point that tracks the hand</param>
        void processPoints(Point3D head, Point3D handRight, Point3D handLeft)
        {
            int deltaHeadRight = head.depth - handRight.depth;
            int deltaHeadLeft = head.depth - handLeft.depth;

            lblDelDepth.Content = "left: " + deltaHeadLeft + "\nright: " + deltaHeadRight + "\nSum: " + (deltaHeadLeft+deltaHeadRight);

            if ((deltaHeadLeft + deltaHeadRight) > (DEPTH_THRESH_TWO))
            {
                //two hands are stretched out
                isHandUp = true;
                lblHandRaised.Content = "Both hands";
                performActionForTwoHands(handLeft, handRight, head);
            }                  
            else 
            {
                nullfiyReference();

                //only one hand is stretched out
                if (deltaHeadRight > DEPTH_THRESH_ONE)
                {
                    //right hand is stretched out
                    isHandUp = true;
                    lblHandRaised.Content = "Right hand";
                    performActionForRightHand(handRight,head);
                }

                else if (deltaHeadLeft > DEPTH_THRESH_ONE)
                {
                    //left hand is stretched out
                    isHandUp = true;
                    lblHandRaised.Content = "Left hand";
                    performActionForLeftHand(handLeft, head);
                    Thread.Sleep((int)(sldDelay.Value+20));
                }

                else
                {
                    lblHandRaised.Content = "Hands down";
                    performActionForNoHands();
                    isHandUp = false;                    
                }
 
            }
            
        }

        /// <summary>
        /// Event handler for skeletal frames. This routine will receive a skeleton frame and track the joints
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void nui_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            SkeletonFrame skFrame = e.SkeletonFrame;
            foreach (SkeletonData skeleton in skFrame.Skeletons)
            {
                //for every skeleton in the frame, check if it is being tracked
                if (skeleton.TrackingState == SkeletonTrackingState.Tracked)
                {
                    //if it is being tracked, read each joint
                    foreach (Joint joint in skeleton.Joints)
                    {
                        //the following lines of code basically convert the coordinate info into the 320x240 space
                        float xCo, yCo;
                        short depthVal;

                        //write the coordinates and depth value of this joint into these local variables
                        nui.SkeletonEngine.SkeletonToDepthImage(joint.Position, out xCo, out yCo, out depthVal);
                        xCo = Math.Max(0, Math.Min(xCo * WIDTH, WIDTH));
                        yCo = Math.Max(0, Math.Min(yCo * HEIGHT, HEIGHT));

                        //if these joints are a wrist or a spine, then save these values into the appropriate class variables
                        switch (joint.ID)
                        {
                            case JointID.Spine:
                                ptHead.x = xCo; ptHead.y = yCo; ptHead.depth = depthVal; ptHead.isValid = true;
                                break;

                            case JointID.WristRight:
                                ptHandRight.x = xCo; ptHandRight.y = yCo; ptHandRight.depth = depthVal; ptHandRight.isValid = true;
                                break;

                            case JointID.WristLeft:
                                ptHandLeft.x = xCo; ptHandLeft.y = yCo; ptHandLeft.depth = depthVal; ptHandLeft.isValid = true;
                                break;
                        }
                                                
                    }

                    //process these points to see if the user's hand is raised
                    processPoints(ptHead, ptHandRight, ptHandLeft);
                }
                else
                {
                    //nothing is being tracked, make these points invalid
                    ptHead.isValid = false;
                    ptHandRight.isValid = false;
                }
                
            }
        }

        /// <summary>
        /// event handler for video frames - this handler basically pulls the color image from the kinect and puts it into the 
        /// appropriate image control. This is done so that the user can see himself and his surroundings in full color 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void nui_VideoFrameReady(object sender, ImageFrameReadyEventArgs e)
        {
            if (chkImageFrame.IsChecked == true)
            {
                //grab the image feed from the sensor
                PlanarImage Image = e.ImageFrame.Image;

                //create the image source from the feed and set it to the source property of the appropriate image box
                imgColor.Source = BitmapSource.Create(Image.Width, Image.Height, 96, 96, PixelFormats.Bgr32, null, Image.Bits, Image.Width * Image.BytesPerPixel);
            }
        }

        /// <summary>
        /// this is the event handler for the depth frame. We simply visualize the depth frame (which is actually a linear array)
        /// as a double dimensional array of pixels where each pixel has a coordinate and a depth value (like a real image). The
        /// depth value and player index is used to calculate the color of the pixel. This is done for every pixel in the depth frame
        /// 
        /// All these color values are then marshalled into a bitmap source that is used to fill the image control in the main window
        /// 
        /// The depth map contains an array of bytes, where 2 bytes represent the info needed for a single pixel. The two bytes contain 
        /// 3 bits for player index (last three bits), and the others are used to store the depth value.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void nui_DepthFrameReady(object sender, ImageFrameReadyEventArgs e)
        {
            byte[] depthFrame = e.ImageFrame.Image.Bits;                                        //this is the depth frame that is given to us
            byte[] imageToShow = new byte[WIDTH*HEIGHT*BPP];                                    //this will store the monochrome image which represents the depth variation using shades of the same color

            int x = 0, y = 0;                                                                   //counters used to facilitate nested loops each representing x and y coords for the pixel
            int indexDepth = 0;                                                                 //the x and y values are used to generate an index that is used to access the linear depth map
            int indexImage = 0;                                                                 //this will keep track of the pixels being added in the image that is used to show the depth map

            const int blue = 0, green = 1, red = 2, size = 3;                                   //constants used to indicate channels for a single pixel. Here the image control uses bgr format

            for (x = 0; x < HEIGHT; x++)
            {
                for (y = 0; y < WIDTH * 2; y += 2, indexImage+=BPP)
                {
                    //(x,y) coordinates will translate to a particular value in the linear array
                    indexDepth = (x * WIDTH * 2) + y;

                    int user = depthFrame[indexDepth] & 0x07;                                           //last three bits give us player index
                    int depth = (depthFrame[indexDepth + 1] << 5) | (depthFrame[indexDepth] >> 3);      //depth value is given by remaining bits
                    byte color = (byte)((depth / MAX_ZONES) + 30);                                      //based on depth value find a suitable shade to represent this depth

                    //user == 0 : indicates part of the scene where there is no user, we'll give it a color based on its depth
                    if (user == 0)
                    {
                        imageToShow[indexImage + blue] = (byte)(color/3);
                        imageToShow[indexImage + green] = (byte)(color / 3);
                        imageToShow[indexImage + red] = (byte)(color / 3);
                    }
                    //any other value for player index, will indicate that the pixel (x,y) is part of the user's body. we'll give that a different color
                    else
                    {
                        imageToShow[indexImage + blue] = 55;
                        imageToShow[indexImage + green] = 55;
                        imageToShow[indexImage + red] = 200;
                    }

                    //the following section uses the values in ptHead and ptHandRightRight (class variables used for tracking joints) and visually represents them in the scene
                    Point3D[] ptsToDraw = {ptHead,ptHandRight,ptHandLeft};

                    //for each point in the list of joints above, compute a bounding box around it, and paint if this pixel at (x,y) is in that box, paint it white
                    foreach(Point3D pt in ptsToDraw)
                    {
                        Point3D currentPixel = new Point3D();
                        currentPixel.x = y / 2;
                        currentPixel.y = x;

                        int bound_left = (int)pt.x - size;
                        int bound_right = (int)pt.x + size;
                        int bound_top = (int)pt.y - size;
                        int bound_bot = (int)pt.y + size;

                        if ((currentPixel.x > bound_left && currentPixel.x < bound_right) && ((currentPixel.y) > bound_top && (currentPixel.y) < bound_bot))
                        {
                            imageToShow[indexImage + blue] = 255;
                            imageToShow[indexImage + green] = 255;
                            imageToShow[indexImage + red] = 255;
                        }
                    }

                }
            }

            //now that the entire depth map has been processed and corresponding color values are stored in the imageToShow array, lets make a bitmap out of this array and display it using image control
            depthMap.Source = BitmapSource.Create((int)depthMap.Width, (int)depthMap.Height, 96, 96, PixelFormats.Bgr24, null, imageToShow, (int)(depthMap.Width * 3));
        }

        /// <summary>
        /// called when the window is closed. Ensure that the keys that were pressed down are released
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closed(object sender, EventArgs e)
        {
            keybd_event((byte)VK_RIGHT, 0, KEYEVENTF_KEYUP, 0);
            keybd_event((byte)VK_LEFT, 0, KEYEVENTF_KEYUP, 0);
        }

        private void btnAngleLess_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                nui.NuiCamera.ElevationAngle -= 10;                
            }
            catch (Exception ex) { }
        }

        private void btnAngleMore_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                nui.NuiCamera.ElevationAngle += 10;                
            }
            catch (Exception ex) { }
        }

        private void chkImageFrame_Checked(object sender, RoutedEventArgs e)
        {
            if (chkImageFrame.IsChecked == true) imgColor.Visibility = Visibility.Visible;
            else imgColor.Visibility = Visibility.Hidden;
        }

        private void sldDelay_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }
    
    }
}
