using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;

namespace N1self.C1selfVisualStudioExtension.ActiveWindow
{
    public class ActiveWindowPoller
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private Timer timer;
        private ForegroundWindowInfo previousActiveWindow;
        private bool isLocked;

        public void Initialise()
        {
            timer = new Timer(_ => { ProcessActiveWindow(); }, null, 0, 10 * 1000);
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            previousActiveWindow = null;

            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                isLocked = true;
                SendForegroundWindowEvent(properties => properties["Title"] = "Session locked");
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                isLocked = false;
                SendForegroundWindowEvent(properties => properties["Title"] = "Session unlocked");
            }
        }

        private void ProcessActiveWindow()
        {
            if (isLocked)
            {
                return;
            }

            var foregroundWindowInfo = GetForegroundWindowInfo();

            if (!Equals(foregroundWindowInfo, previousActiveWindow))
            {
                SendForegroundWindowEvent(properties =>
                {
                    if (!string.IsNullOrWhiteSpace(foregroundWindowInfo.Path))
                    {
                        properties["Process"] = foregroundWindowInfo.Path;
                    }

                    properties["Title"] = foregroundWindowInfo.WindowTitle;
                });

                previousActiveWindow = foregroundWindowInfo;
            }
        }

        private ForegroundWindowInfo GetForegroundWindowInfo()
        {
            var foregroundWindow = GetForegroundWindow();
            var windowTextLength = GetWindowTextLength(foregroundWindow) + 1;
            var windowTitle = new StringBuilder(windowTextLength);

            if (windowTextLength > 0)
            {
                GetWindowText(foregroundWindow, windowTitle, windowTextLength);
            }

            var processPath = GetProcessPathFromWindowHandle(foregroundWindow);

            return new ForegroundWindowInfo
            {
                Path = processPath,
                WindowTitle = windowTitle.ToString()
            };
        }

        private string GetProcessPathFromWindowHandle(IntPtr hwnd)
        {
            try
            {
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                var p = Process.GetProcessById((int) pid);
                return p.MainModule.FileName;
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
                return null;
            }
        }

        private void SendForegroundWindowEvent(Action<JObject> setPropertiesCallback)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", Settings.Default.WriteToken);

            var activityEvent = new JObject();
            activityEvent["dateTime"] = DateTime.Now.ToString("o");

            var location = new JObject();
            location["lat"] = Settings.Default.Latitude;
            location["long"] = Settings.Default.Longitude;
            activityEvent["location"] = location;

            activityEvent["actionTags"] = new JArray("Sample");
            activityEvent["objectTags"] = new JArray("Computer", "ActiveWindow");

            JObject properties = new JObject();
            setPropertiesCallback(properties);
            activityEvent["properties"] = properties;

            var url = string.Format("https://api.1self.co/v1/streams/{0}/events", Settings.Default.StreamId);
            var content = new StringContent(activityEvent.ToString(Newtonsoft.Json.Formatting.None));
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            Debug.WriteLine(activityEvent.ToString());

            client.PostAsync(url, content).ContinueWith(postTask =>
            {
                try
                {
                    Debug.WriteLine(postTask.Result.StatusCode.ToString(), "1SELF");
                }
                catch (Exception)
                {
                    Debug.WriteLine("1self: Couldn't send current active window event");
                }
            });
        }
    }
}