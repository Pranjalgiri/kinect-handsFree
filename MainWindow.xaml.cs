using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Kinect;
using Microsoft.Speech.Recognition;
using System.Threading;
using Microsoft.Speech.AudioFormat;
using System.Diagnostics;
using System.Timers;
using System.Runtime.InteropServices;
using System.Drawing;
using PowerPoint_kinect.Gestures;


namespace PowerPoint_kinect
{
    public partial class MainWindow : Window
    {
        bool mouse = false;
        bool isDragOn = false;
        bool isLengthCallibrated = false;
        
        double[] x = new double[3];
        double[] y = new double[3];
        double originX, originY;

        double armLength;

        int framecount, count;

        KinectSensor sensor;
        
        SpeechRecognitionEngine speechRecognizer = new SpeechRecognitionEngine();

        [DllImport("user32")]
        public static extern int SetCursorPos(int x, int y);
        
        private const int MOUSEEVENTF_MOVE = 0x01;

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;

        private const int MOUSEEVENTF_LEFTUP = 0x04;

        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;

        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        private const int MOUSEEVENTF_MIDDLEDOWN = 0x20;

        private const int MOUSEEVENTF_MIDDLEUP = 0x40;

        private const int MOUSEEVENTF_ABSOLUTE = 0x80;

        [DllImport("user32.dll",
            CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]

        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        byte[] colorBytes;
        
        Skeleton[] skeletons;

        Skeleton closestSkeleton;
        
        bool isCirclesVisible = true;

        private GestureController gestureController;

        public MainWindow()
        {
            
            InitializeComponent();

            framecount = 0;

            count = 0;

            armLength = 0;

            this.Loaded += new RoutedEventHandler(MainWindow_Loaded);

            
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }
        
        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            sensor = KinectSensor.KinectSensors.FirstOrDefault();
            if (sensor == null)
            {
                MessageBox.Show("This application requires a Kinect sensor.");
                this.Close();
            }

            sensor.Start();

            sensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            sensor.ColorFrameReady += new EventHandler<ColorImageFrameReadyEventArgs>(sensor_ColorFrameReady);

            sensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
            TransformSmoothParameters smoothingParam = new TransformSmoothParameters();
            {
                smoothingParam.Smoothing = 0.5f;
                smoothingParam.Correction = 0.5f;
                smoothingParam.Prediction = 0.5f;
                smoothingParam.JitterRadius = 0.1f;
                smoothingParam.MaxDeviationRadius = 0.1f;
            }

            sensor.SkeletonStream.Enable(smoothingParam); // Enable skeletal tracking

            sensor.SkeletonStream.Enable();
            if ((bool)cb1.IsChecked)
            {
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                sensor.SkeletonStream.EnableTrackingInNearRange = true;
            }else
            {
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
            }

            framecount++;

            sensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(sensor_SkeletonFrameReady);

            // initialize the gesture recognizer
            gestureController = new GestureController();
            gestureController.GestureRecognized += OnGestureRecognized;

            // register the gestures for this demo
            RegisterGestures();

            Application.Current.Exit += new ExitEventHandler(Current_Exit);

            //Thread speechThread = new Thread(new ThreadStart(InitializeSpeechRecognition)); 
            InitializeSpeechRecognition();
        }
        
        void Current_Exit(object sender, ExitEventArgs e)
        {
            if (speechRecognizer != null)
            {
                speechRecognizer.RecognizeAsyncCancel();
                speechRecognizer.RecognizeAsyncStop();
            }
            if (sensor != null)
            {
                sensor.AudioSource.Stop();
                sensor.Stop();
                sensor.Dispose();
                sensor = null;
            }
        }
        
        void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.C)
            {
                ToggleCircles();
            }
        }
        
        void sensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (var image = e.OpenColorImageFrame())
            {
                if (image == null)
                    return;

                if (colorBytes == null ||
                    colorBytes.Length != image.PixelDataLength)
                {
                    colorBytes = new byte[image.PixelDataLength];
                }

                image.CopyPixelDataTo(colorBytes);

                
                int length = colorBytes.Length;
                for (int i = 0; i < length; i += 4)
                {
                    colorBytes[i + 3] = 255;
                }

                BitmapSource source = BitmapSource.Create(image.Width,
                    image.Height,
                    96,
                    96,
                    PixelFormats.Bgra32,
                    null,
                    colorBytes,
                    image.Width * image.BytesPerPixel);
                videoImage.Source = source;
            }
        }
        
        void sensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (var skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame == null)
                    return;

                if (skeletons == null ||
                    skeletons.Length != skeletonFrame.SkeletonArrayLength)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                }

                skeletonFrame.CopySkeletonDataTo(skeletons);
            }

            closestSkeleton = skeletons.Where(s => s.TrackingState == SkeletonTrackingState.Tracked)
                                                .OrderBy(s => s.Position.Z * Math.Abs(s.Position.X))
                                                .FirstOrDefault();

            if (closestSkeleton == null)
                return;

            Joint head = closestSkeleton.Joints[JointType.Head];
            Joint rightHand = closestSkeleton.Joints[JointType.HandRight];
            Joint leftHand = closestSkeleton.Joints[JointType.HandLeft];
            Joint rightElbow = closestSkeleton.Joints[JointType.ElbowRight];
            Joint leftElbow = closestSkeleton.Joints[JointType.ElbowLeft];
            Joint rightShoulder = closestSkeleton.Joints[JointType.ShoulderRight];
            Joint leftShoulder = closestSkeleton.Joints[JointType.ShoulderLeft];
            Joint centerShoulder = closestSkeleton.Joints[JointType.ShoulderCenter];
            Joint spine = closestSkeleton.Joints[JointType.Spine];
            Joint centerHip = closestSkeleton.Joints[JointType.HipCenter];

            SetEllipsePosition(ellipseHead, head);
            SetEllipsePosition(ellipseLeftHand, leftHand);
            SetEllipsePosition(ellipseRightHand, rightHand);
            SetEllipsePosition(ellipseLeftElbow, leftElbow);
            SetEllipsePosition(ellipseRightElbow, rightElbow);
            SetEllipsePosition(ellipseLeftShoulder, leftShoulder);
            SetEllipsePosition(ellipseRightShoulder, rightShoulder);
            SetEllipsePosition(ellipseShoulderCenter, centerShoulder);
            SetEllipsePosition(ellipseSpine, spine);
            SetEllipsePosition(ellipseHipCenter, centerHip);

            if (!isLengthCallibrated) l1.Content = "Callibrating";

            if(Math.Abs(leftHand.Position.X - leftElbow.Position.X) < 0.1 && Math.Abs(leftHand.Position.Z - centerShoulder.Position.Z + 0.2) < 0.1 && count != 20 && !isLengthCallibrated && Math.Abs(leftHand.Position.Y - centerShoulder.Position.Y) > 0.3)
            {
                armLength = (armLength*count + Math.Abs(leftHand.Position.Y - leftElbow.Position.Y))/(count+1);
                count++;
                return;
            }
            else if(count == 20)
            {
                isLengthCallibrated = true;
                armLength = armLength + 0.1;
                l1.Content = "Ready. ArmLength = " + armLength;
                count++;
            }

            if (mouse)
            {
                double i, j;
                i = 0.4 * rightHand.Position.X + 0.3 * x[0] + 0.2 * x[1] + 0.1 * x[2];
                j = 0.4 * rightHand.Position.Y + 0.3 * y[0] + 0.2 * y[1] + y[2] * 0.1;
                for (int t = 0; t < 2; t++)
                {
                    x[t + 1] = x[t];
                    y[t + 1] = y[t];
                }
                x[0] = rightHand.Position.X;
                y[0] = rightHand.Position.Y;
                int a = (int)Math.Floor((i - originX - 0.05) * 3000);
                int b = 360 + (int)Math.Floor((j - originY + .1) * -2000);
                SetCursorPos(a, b);
                Thread.Sleep(10);
                if((leftHand.Position.Z < centerShoulder.Position.Z - 0.5) && !isDragOn)
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                    isDragOn = true;
                } else if (!(leftHand.Position.Z < centerShoulder.Position.Z - 0.40) && isDragOn)
                {
                    mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                    isDragOn = false;
                }
                //SystemCursors.setCustomCursor("C:\\Users\\Rituj Beniwal\\Desktop\\PowerPoint_kinect\\PowerPoint_kinect\\cursor.cur");
            }

            if (head.TrackingState == JointTrackingState.NotTracked ||
                rightHand.TrackingState == JointTrackingState.NotTracked ||
                leftHand.TrackingState == JointTrackingState.NotTracked)
            {
                //Don't have a good read on the joints so we cannot process gestures
                return;
            }

            // update the gesture controller
            gestureController.UpdateAllGestures(closestSkeleton);
        }

        private void RegisterGestures()
        {

            IRelativeGestureSegment[] swipeleftSegments = new IRelativeGestureSegment[3];
            swipeleftSegments[0] = new SwipeLeftSegment1();
            swipeleftSegments[1] = new SwipeLeftSegment2();
            swipeleftSegments[2] = new SwipeLeftSegment3();
            gestureController.AddGesture("SwipeLeft", swipeleftSegments);

            IRelativeGestureSegment[] swiperightSegments = new IRelativeGestureSegment[3];
            swiperightSegments[0] = new SwipeRightSegment1();
            swiperightSegments[1] = new SwipeRightSegment2();
            swiperightSegments[2] = new SwipeRightSegment3();
            gestureController.AddGesture("SwipeRight", swiperightSegments);

            IRelativeGestureSegment[] zoomInSegments = new IRelativeGestureSegment[3];
            zoomInSegments[0] = new ZoomSegment1();
            zoomInSegments[1] = new ZoomSegment2();
            zoomInSegments[2] = new ZoomSegment3();
            gestureController.AddGesture("ZoomIn", zoomInSegments);

            IRelativeGestureSegment[] zoomOutSegments = new IRelativeGestureSegment[3];
            zoomOutSegments[0] = new ZoomSegment3();
            zoomOutSegments[1] = new ZoomSegment2();
            zoomOutSegments[2] = new ZoomSegment1();
            gestureController.AddGesture("ZoomOut", zoomOutSegments);

            IRelativeGestureSegment[] swipeUpSegments = new IRelativeGestureSegment[3];
            swipeUpSegments[0] = new SwipeVerticalSegment1();
            swipeUpSegments[1] = new SwipeVerticalSegment2();
            swipeUpSegments[2] = new SwipeVerticalSegment3();
            gestureController.AddGesture("SwipeUp", swipeUpSegments);

            IRelativeGestureSegment[] swipeDownSegments = new IRelativeGestureSegment[3];
            swipeDownSegments[0] = new SwipeVerticalSegment3();
            swipeDownSegments[1] = new SwipeVerticalSegment2();
            swipeDownSegments[2] = new SwipeVerticalSegment1();
            gestureController.AddGesture("SwipeDown", swipeDownSegments);

            IRelativeGestureSegment[] hideWindowSegments = new IRelativeGestureSegment[3];
            hideWindowSegments[0] = new HideWindowSegment1();
            hideWindowSegments[1] = new HideWindowSegment2();
            hideWindowSegments[2] = new HideWindowSegment3();
            gestureController.AddGesture("HideWindow", hideWindowSegments);

            IRelativeGestureSegment[] showWindowSegments = new IRelativeGestureSegment[3];
            showWindowSegments[0] = new HideWindowSegment3();
            showWindowSegments[1] = new HideWindowSegment2();
            showWindowSegments[2] = new HideWindowSegment1();
            gestureController.AddGesture("ShowWindow", showWindowSegments);

            IRelativeGestureSegment[] waveRightSegments = new IRelativeGestureSegment[20];
            for(int i = 0; i < 20; i = i+2)
            {
                if (i % 4 == 0)
                {
                    for (int j = i; j < i+2; j++)
                    {
                        waveRightSegments[j] = new WaveRightSegment1();
                    }
                }
                else
                {
                    for (int j = i; j < i+2; j++)
                    {
                        waveRightSegments[j] = new WaveRightSegment2();
                    }
                }
            }
            gestureController.AddGesture("WaveRight", waveRightSegments);

        }



        /// <summary>
        /// Gets or sets the last recognized gesture.
        /// </summary>
        private string _gesture;
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        public String Gesture
        {
            get { return _gesture; }

            private set
            {
                if (_gesture == value)
                    return;

                _gesture = value;

                if (this.PropertyChanged != null) 
                    PropertyChanged(this, new System.ComponentModel.PropertyChangedEventArgs("Gesture"));
            }
        }

        private void OnGestureRecognized(object sender, GestureEventArgs e)
        {
            if (!mouse)
            {
                switch (e.GestureName)
                {
                    case "WaveRight":
                        Gesture = "Wave Right";
                        l1.Content = "Wave Right";
                        if (closestSkeleton.TrackingState == SkeletonTrackingState.Tracked)
                        {
                            mouse = true;
                            originX = closestSkeleton.Joints[JointType.ShoulderCenter].Position.X;
                            originY = closestSkeleton.Joints[JointType.ShoulderCenter].Position.Y;
                            x[0] = x[1] = x[2] = closestSkeleton.Joints[JointType.HandRight].Position.X;
                            y[0] = y[1] = y[2] = closestSkeleton.Joints[JointType.HandRight].Position.Y;
                            Trace.WriteLine("Mouse Mode ON");
                        }
                        else
                        {
                            mouse = false;
                            Trace.WriteLine("No skeleton found.");
                            MessageBox.Show("No skeleton found. Are you there?");
                        }
                        break;
                    case "SwipeLeft":
                        Gesture = "Swipe Left";
                        l1.Content = "Swipe Left";
                        System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                        break;
                    case "SwipeRight":
                        Gesture = "Swipe Right";
                        l1.Content = "Swipe Right";
                        System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                        break;
                    case "SwipeUp":
                        Gesture = "Swipe Up";
                        l1.Content = "Swipe Up";
                        System.Windows.Forms.SendKeys.SendWait("{DOWN 4}");
                        break;
                    case "SwipeDown":
                        Gesture = "Swipe Down";
                        l1.Content = "Swipe Down";
                        System.Windows.Forms.SendKeys.SendWait("{UP 3}");
                        break;
                    case "ZoomIn":
                        Gesture = "Zoom In";
                        l1.Content = "Zoom In";
                        System.Windows.Forms.SendKeys.SendWait("^{ADD}");
                        break;
                    case "ZoomOut":
                        Gesture = "Zoom Out";
                        l1.Content = "Zoom Out";
                        System.Windows.Forms.SendKeys.SendWait("^{SUBTRACT}");
                        break;
                    case "HideWindow":
                        Gesture = "Hide Window";
                        l1.Content = "Hide Window";
                        this.Dispatcher.BeginInvoke((Action)delegate
                        {
                            HideWindow();
                        });
                        break;
                    case "ShowWindow":
                        Gesture = "Show Window";
                        l1.Content = "Show Window";
                        this.Dispatcher.BeginInvoke((Action)delegate
                        {
                            ShowWindow();
                        });
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Clear text after some time
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void clearTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Gesture = "";
            //_clearTimer.Stop();
        }

        private void SetEllipsePosition(Ellipse ellipse, Joint joint)
        {
            ellipse.Width = 10;
            ellipse.Height = 10;
            ellipse.Fill = new SolidColorBrush(Colors.Red);

            CoordinateMapper mapper = sensor.CoordinateMapper;
            
            var point = mapper.MapSkeletonPointToColorPoint(joint.Position, sensor.ColorStream.Format);

            Canvas.SetLeft(ellipse, point.X - ellipse.ActualWidth / 2);
            Canvas.SetTop(ellipse, point.Y - ellipse.ActualHeight / 2);
        }
        
        void ToggleCircles()
        {
            if (isCirclesVisible)
                HideCircles();
            else
                ShowCircles();
        }
        
        void HideCircles()
        {
            isCirclesVisible = false;
            ellipseHead.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightHand.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftElbow.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightElbow.Visibility = System.Windows.Visibility.Collapsed;
            ellipseLeftShoulder.Visibility = System.Windows.Visibility.Collapsed;
            ellipseRightShoulder.Visibility = System.Windows.Visibility.Collapsed;
        }
        
        void ShowCircles()
        {
            isCirclesVisible = true;
            ellipseHead.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftHand.Visibility = System.Windows.Visibility.Visible;
            ellipseRightHand.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftElbow.Visibility = System.Windows.Visibility.Visible;
            ellipseRightElbow.Visibility = System.Windows.Visibility.Visible;
            ellipseLeftShoulder.Visibility = System.Windows.Visibility.Visible;
            ellipseRightShoulder.Visibility = System.Windows.Visibility.Visible;
        }
        
        private void ShowWindow()
        {
            this.Topmost = true;
            this.WindowState = System.Windows.WindowState.Maximized;
        }
        
        private void HideWindow()
        {
            this.Topmost = false;
            this.WindowState = System.Windows.WindowState.Minimized;
        }

        #region Speech Recognition Methods

        private static RecognizerInfo GetKinectRecognizer()
        {
            Func<RecognizerInfo, bool> matchingFunc = r =>
            {
                string value;
                r.AdditionalInfo.TryGetValue("Kinect", out value);
                return "True".Equals(value, StringComparison.InvariantCultureIgnoreCase) && "en-US".Equals(r.Culture.Name, StringComparison.InvariantCultureIgnoreCase);
            };
            return SpeechRecognitionEngine.InstalledRecognizers().Where(matchingFunc).FirstOrDefault();
        }

        private void InitializeSpeechRecognition()
        {
            Choices words = new Choices();
            words.Add(new string[] { "previous slide", "next slide", "hide window", "show window", "Backspace","close window", "drag", "drop", "Start seated mode", "Start near mode", "Stop seated mode", "Stop near mode", "Mouse on", "Mouse off", "left", "right", "open", "Close Application", "Page Up", "Page Down", "Undo", " Zoom In", "Zoom Out", "Delete Permanently", "up", "Down", "Enter", "Escape", "Open Task Manager"});
            GrammarBuilder gb = new GrammarBuilder();
            gb.Append(words);

            Grammar g = new Grammar(gb);
            speechRecognizer.LoadGrammar(g);

            speechRecognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);

            StartSpeechRecognition();

        }

        private void StartSpeechRecognition()
        {
            if (sensor == null || speechRecognizer == null)
                return;

            var audioSource = this.sensor.AudioSource;
            audioSource.BeamAngleMode = BeamAngleMode.Adaptive;
            var kinectStream = audioSource.Start();

            speechRecognizer.SetInputToDefaultAudioDevice();
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple);

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            count = 0;
            isLengthCallibrated = false;
            Thread.Sleep(100);
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            if (e.Result.Confidence < 0.30)
            {
                Trace.WriteLine("\nSpeech Rejected filtered, confidence: " + e.Result.Confidence);
                return;
            }

            Trace.WriteLine("\nSpeech Recognized, confidence: " + e.Result.Confidence + ": \t{0}", e.Result.Text);

            if (e.Result.Text == "show window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    ShowWindow();
                });
            }
            else if (e.Result.Text == "hide window")
            {
                this.Dispatcher.BeginInvoke((Action)delegate
                {
                    HideWindow();
                });
            }else if (e.Result.Text == "next slide")
            {
                System.Windows.Forms.SendKeys.SendWait("{Right}");
            }else if (e.Result.Text == "previous slide")
            {
                System.Windows.Forms.SendKeys.SendWait("{Left}");
            }
            else if (e.Result.Text == "close window")
            {
                this.Close();
            }else if(e.Result.Text == "Start seated mode" || e.Result.Text == "Start near mode")
            {
                cb1.IsChecked = true;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                sensor.SkeletonStream.EnableTrackingInNearRange = true;
            }
            else if (e.Result.Text == "Stop seated mode" || e.Result.Text == "Stop near mode")
            {
                cb1.IsChecked = false;
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
            }
            else if (e.Result.Text == "Mouse on")
            {
                if (closestSkeleton.TrackingState == SkeletonTrackingState.Tracked)
                {
                    mouse = true;
                    originX = closestSkeleton.Joints[JointType.ShoulderCenter].Position.X;
                    originY = closestSkeleton.Joints[JointType.ShoulderCenter].Position.Y;
                    x[0] = x[1] = x[2] = closestSkeleton.Joints[JointType.HandRight].Position.X;
                    y[0] = y[1] = y[2] = closestSkeleton.Joints[JointType.HandRight].Position.Y;
                    Trace.WriteLine("Mouse Mode ON");
                }else
                {
                    mouse = false;
                    Trace.WriteLine("No skeleton found.");
                    MessageBox.Show("No skeleton found. Are you there?");
                }
            }
            else if (e.Result.Text == "Mouse off")
            {
                mouse = false;
                Trace.WriteLine("Mouse Mode OFF");
            }
            else if (e.Result.Text == "drag")
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
            }
            else if (e.Result.Text == "drop")
            {
                mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
            }
            else if (e.Result.Text == "left")
            {
               
                mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                
            }
            else if (e.Result.Text == "right")
            {
                mouse_event(MOUSEEVENTF_RIGHTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                
                mouse_event(MOUSEEVENTF_RIGHTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
            }
            else if (e.Result.Text == "open")
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                Thread.Sleep(100);
                mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);


            }
            else if (e.Result.Text == "Close Application")
            {
                System.Windows.Forms.SendKeys.SendWait("%{F4}");
            }
            else if (e.Result.Text == " Page Up")
            {
                System.Windows.Forms.SendKeys.SendWait("{PGUP}");
            }
            else if (e.Result.Text == "Page Down")
            {
                System.Windows.Forms.SendKeys.SendWait("{PGDN}");
            }
            else if (e.Result.Text == "Undo")
            {
                System.Windows.Forms.SendKeys.SendWait("^{Z}");
            }
            else if (e.Result.Text == "Zoom In")
            {
                System.Windows.Forms.SendKeys.SendWait("^{ADD}");
            }
            else if (e.Result.Text == "Zoom Out")
            {
                System.Windows.Forms.SendKeys.SendWait("^{SUBTRACT}");
            }
            else if (e.Result.Text == "Backspace")
            {
                System.Windows.Forms.SendKeys.SendWait("{BS}");
            }
            else if (e.Result.Text == "Delete Permanently")
            {
                System.Windows.Forms.SendKeys.SendWait("+{DEL}");
            }
            else if (e.Result.Text == "Up")
            {
                System.Windows.Forms.SendKeys.SendWait("{UP}");
            }
            else if (e.Result.Text == "Down")
            {
                System.Windows.Forms.SendKeys.SendWait("{DOWN}");
            }
            else if (e.Result.Text == "Enter")
            {
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            }
            else if (e.Result.Text == "Escape")
            {
                System.Windows.Forms.SendKeys.SendWait("{ESC}");
            }
            else if (e.Result.Text == "Open Task Manager")
            {
                System.Windows.Forms.SendKeys.SendWait("^+{ESC}");
            }
        }

        #endregion

        private void cb1_Checked(object sender, RoutedEventArgs e)
        {
            if ((bool)cb1.IsChecked)
            {
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated;
                sensor.SkeletonStream.EnableTrackingInNearRange = true;
            }else
            {
                sensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                sensor.SkeletonStream.EnableTrackingInNearRange = false;
            }
        }
    }

}
