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

        private Timer timer;
        private string previousActiveWindowTitle;

        public void Initialise()
        {
            timer = new Timer(_ => { ProcessActiveWindow(); }, null, 0, 10 * 1000);
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            previousActiveWindowTitle = null;

            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                SendForegroundWindowEvent(properties => properties["Title"] = "Session locked");
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                SendForegroundWindowEvent(properties => properties["Title"] = "Session unlocked");
            }
        }

        private void ProcessActiveWindow()
        {
            string windowTitle = GetForegroundWindowTitle();

            if (!string.IsNullOrWhiteSpace(windowTitle)
                && !string.Equals(previousActiveWindowTitle, windowTitle))
            {
                SendForegroundWindowEvent(properties => properties["Title"] = windowTitle);
                previousActiveWindowTitle = windowTitle;
            }
        }

        private string GetForegroundWindowTitle()
        {
            var foregroundWindow = GetForegroundWindow();
            var windowTextLength = GetWindowTextLength(foregroundWindow) + 1;

            if (windowTextLength > 0)
            {
                var windowTitle = new StringBuilder(windowTextLength);

                if (GetWindowText(foregroundWindow, windowTitle, windowTextLength) > 0)
                {
                    return windowTitle.ToString();
                }
            }

            return null;
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