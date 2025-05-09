using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Text.Json;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Linq;
using System.Collections.Generic;

namespace BibleVerseOpener
{
    class Program
    {
        // Win32 API Funktionen
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // Konstanten
        private const int MOD_CONTROL = 0x0002;
        private const int VK_C = 0x43;
        private const int HOTKEY_ID = 1;

        // Konfigurationsklasse
        public class AppConfig
        {
            public string[] SupportedLanguages { get; set; }
            public Dictionary<string, string> BibleUrls { get; set; }
            public Dictionary<string, Dictionary<string, string>> BookMappings { get; set; }
        }

        // Statische Konfigurationsvariablen
        private static AppConfig config;
        private static Regex BibleVerseRegex;

        private static DateTime lastCopyTime = DateTime.MinValue;
        private static NotifyIcon notifyIcon;

        // Variablen für Debouncing
        private static DateTime lastClipboardEvent = DateTime.MinValue;
        private static readonly TimeSpan debounceDuration = TimeSpan.FromMilliseconds(100);

        [STAThread]
        static void Main(string[] args)
        {
            // Lade Konfiguration
            LoadConfiguration();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Erstelle den Icon für den Tray
            Icon bibleIcon = CreateBibleIcon();

            // Erstelle eine NotifyIcon für den Tray
            notifyIcon = new NotifyIcon
            {
                Icon = bibleIcon,
                Text = "Bibelvers-Opener",
                Visible = true
            };

            // Erstelle ein Kontextmenü für den Tray-Icon
            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Beenden", null, (s, e) => Application.Exit());
            contextMenu.Items.Add("Info", null, ShowInfo);
            notifyIcon.ContextMenuStrip = contextMenu;

            // Doppelklick auf Icon zeigt Info an
            notifyIcon.DoubleClick += (s, e) => ShowInfo(null, null);

            // Registriere den Hotkey
            HookClipboard();

            // Zeige eine Benachrichtigung, dass die App gestartet wurde
            notifyIcon.ShowBalloonTip(3000, "Bibelvers-Opener",
                $"Anwendung läuft im Hintergrund. " +
                $"Drücke 2x STRG+C innerhalb von 1 Sekunden, um einen Bibelvers zu öffnen.",
                ToolTipIcon.Info);

            // Halte die Anwendung am Laufen
            Application.Run();

            // Aufräumen beim Beenden
            notifyIcon.Visible = false;
            UnhookClipboard();
        }

        private static void LoadConfiguration()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

            try
            {
                string jsonString = File.ReadAllText(configPath);
                config = JsonSerializer.Deserialize<AppConfig>(jsonString, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (config == null)
                {
                    throw new Exception("Deserialisierung ergab ein null-Objekt.");
                }

                // Dynamische Regex-Generierung basierend auf Konfiguration
                BuildBibleVerseRegex();
            }
            catch (JsonException jsonEx)
            {
                MessageBox.Show($"JSON-Fehler beim Laden der Konfiguration: {jsonEx.Message}",
                    "Konfigurationsfehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Laden der Konfiguration: {ex.Message}",
                    "Konfigurationsfehler", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private static void BuildBibleVerseRegex()
        {
            // Baue die Regex dynamisch aus den Buchmappings aller unterstützten Sprachen
            var allBookPatterns = config.SupportedLanguages.SelectMany(lang =>
                config.BookMappings[lang].Keys.Concat(config.BookMappings[lang].Values)
            );

            string bookPattern = string.Join("|", allBookPatterns);

            BibleVerseRegex = new Regex(
                @$"({bookPattern})\.?\s*(\d+)(?:\s*,\s*(\d+)(?:\s*-\s*(\d+))?)?",
                RegexOptions.Compiled | RegexOptions.IgnoreCase
            );
        }

        private static void ShowInfo(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Bibelvers-Opener\n\n" +
                "Markiere einen Bibelvers und drücke zweimal STRG+C innerhalb von 1 Sekunden.\n" +
                "Das Programm erkennt den Vers und öffnet ihn in deinem Browser.\n\n" +
                "Unterstützte Formate:\n" +
                "- 1Mo 3,16\n" +
                "- Gen 3,16\n" +
                "- 1Mo 3,16-18",
                "Über Bibelvers-Opener",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        private static void ProcessBibleVerses(string text)
        {
            var matches = BibleVerseRegex.Matches(text);

            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Extrahiere die gefundenen Gruppen
                    string bookAbbr = match.Groups[1].Value.Trim();
                    string chapter = match.Groups[2].Value.Trim();
                    string verse = match.Groups[3].Value.Trim();
                    string endVerse = match.Groups[4].Value.Trim();

                    // Finde die Sprache und den vollständigen Buchnamen
                    string detectedLanguage = null;
                    string fullBookName = null;

                    foreach (var lang in config.SupportedLanguages)
                    {
                        if (config.BookMappings[lang].ContainsKey(bookAbbr))
                        {
                            detectedLanguage = lang;
                            fullBookName = config.BookMappings[lang][bookAbbr];
                            break;
                        }
                        else if (config.BookMappings[lang].ContainsValue(bookAbbr))
                        {
                            detectedLanguage = lang;
                            fullBookName = bookAbbr;
                            break;
                        }
                    }

                    if (detectedLanguage == null)
                    {
                        // Keine passende Sprache gefunden, überspringen
                        continue;
                    }

                    // Baue den Link auf
                    string url = $"{config.BibleUrls[detectedLanguage]}{Uri.EscapeDataString(fullBookName)}{chapter}";

                    if (!string.IsNullOrEmpty(verse))
                    {
                        url += $",{verse}";
                        if (!string.IsNullOrEmpty(endVerse))
                        {
                            url += $"-{endVerse}";
                        }
                    }

                    // Öffne den Link im Standardbrowser
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    });

                    // Zeige eine Benachrichtigung
                    //notifyIcon.ShowBalloonTip(2000, "Bibelvers gefunden",
                    //    $"Öffne {fullBookName} {chapter}{(!string.IsNullOrEmpty(verse) ? $",{verse}" : "")}" +
                    //    $"{(!string.IsNullOrEmpty(endVerse) ? $"-{endVerse}" : "")}",
                    //    ToolTipIcon.Info);

                    // Nur den ersten gefundenen Vers verarbeiten
                    break;
                }
            }
            else
            {
                //notifyIcon.ShowBalloonTip(2000, "Kein Bibelvers gefunden",
                //    "Im kopierten Text wurde kein Bibelvers erkannt.", ToolTipIcon.Warning);
            }
        }

        private static void HookClipboard()
        {
            // Registriere den EventHandler für die Zwischenablage
            ClipboardNotification.ClipboardChanged += OnClipboardChanged;
        }

        private static void UnhookClipboard()
        {
            // Entferne den EventHandler
            ClipboardNotification.ClipboardChanged -= OnClipboardChanged;
        }

        private static bool isWaitingForSecondCopy = false;
        private static System.Threading.Timer resetTimer;
        private static string firstCopyContent = "";

        private static void OnClipboardChanged(object sender, EventArgs e)
        {
            try
            {
                // Implementiere Debouncing: Ignoriere Events, die zu schnell aufeinander folgen
                DateTime now = DateTime.Now;
                if (now - lastClipboardEvent < debounceDuration)
                {
                    // Event ignorieren, da es zu schnell nach dem letzten kommt
                    return;
                }

                // Aktualisiere den Zeitstempel für das Debouncing
                lastClipboardEvent = now;

                // Verarbeite den Clipboard-Inhalt
                string clipboardText = Clipboard.GetText();

                // Beim ersten STRG+C wird der Timer gestartet
                if (!isWaitingForSecondCopy)
                {
                    lastCopyTime = now;
                    isWaitingForSecondCopy = true;
                    firstCopyContent = clipboardText;

                    // Timer zum Zurücksetzen des Status, wenn kein zweiter STRG+C innerhalb von 1 Sekunden erfolgt
                    if (resetTimer == null)
                    {
                        resetTimer = new System.Threading.Timer((state) =>
                        {
                            isWaitingForSecondCopy = false;
                        }, null, 1000, Timeout.Infinite);
                    }
                    else
                    {
                        resetTimer.Change(1000, Timeout.Infinite);
                    }

                    // Optional: Zeige dem Benutzer einen Hinweis an, dass der erste STRG+C erkannt wurde
                    //notifyIcon.ShowBalloonTip(1000, "Bereit",
                    //    "Erster STRG+C erkannt. Drücke nochmal für Bibelvers-Suche.",
                    //    ToolTipIcon.Info);
                }
                // Beim zweiten STRG+C wird der Inhalt verarbeitet
                else if ((now - lastCopyTime).TotalSeconds <= 1)
                {
                    // Setze den Status zurück
                    isWaitingForSecondCopy = false;
                    resetTimer.Change(Timeout.Infinite, Timeout.Infinite);

                    if (clipboardText == firstCopyContent)
                    {
                        if (!string.IsNullOrWhiteSpace(clipboardText))
                        {
                            ProcessBibleVerses(clipboardText);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //notifyIcon.ShowBalloonTip(3000, "Fehler",
                //    $"Fehler beim Verarbeiten der Zwischenablage: {ex.Message}",
                //    ToolTipIcon.Error);
            }
        }

        private static Icon CreateBibleIcon()
        {
            try
            {
                // Erstelle ein Bitmap für das Icon (32x32 Pixel)
                Bitmap bitmap = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Einstellungen für hohe Qualität
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    g.Clear(Color.Transparent);

                    // Buch zeichnen (Hintergrund)
                    using (GraphicsPath bookPath = new GraphicsPath())
                    {
                        // Bucheinband
                        Rectangle bookRect = new Rectangle(2, 5, 28, 24);
                        bookPath.AddRectangle(bookRect);

                        // Fülle das Buch mit einem Farbverlauf
                        using (LinearGradientBrush bookBrush = new LinearGradientBrush(
                            bookRect, Color.DarkRed, Color.Firebrick, LinearGradientMode.Vertical))
                        {
                            g.FillPath(bookBrush, bookPath);
                        }

                        // Rand des Buches
                        using (Pen bookPen = new Pen(Color.Maroon, 1.5f))
                        {
                            g.DrawPath(bookPen, bookPath);
                        }
                    }

                    // Buchseiten
                    using (Brush pagesBrush = new SolidBrush(Color.AntiqueWhite))
                    {
                        g.FillRectangle(pagesBrush, new Rectangle(5, 8, 22, 18));
                    }

                    // Bücherrücken
                    using (Pen bookSpinePen = new Pen(Color.DarkRed, 1f))
                    {
                        g.DrawLine(bookSpinePen, new Point(4, 5), new Point(4, 29));
                    }

                    // Text "B" für Bibel
                    using (Font font = new Font("Arial", 12, FontStyle.Bold))
                    using (Brush textBrush = new SolidBrush(Color.Navy))
                    {
                        g.DrawString("B", font, textBrush, new PointF(10, 7));
                    }

                    // Kreuz-Symbol hinzufügen
                    using (Pen crossPen = new Pen(Color.DarkGoldenrod, 2f))
                    {
                        g.DrawLine(crossPen, new Point(22, 10), new Point(22, 24));
                        g.DrawLine(crossPen, new Point(17, 15), new Point(27, 15));
                    }
                }

                // Konvertiere das Bitmap in ein Icon
                IntPtr hIcon = bitmap.GetHicon();
                Icon icon = Icon.FromHandle(hIcon);

                // Erstelle eine Kopie des Icons (um Speicherlecks zu vermeiden)
                using (MemoryStream ms = new MemoryStream())
                {
                    icon.Save(ms);
                    ms.Position = 0;
                    return new Icon(ms);
                }
            }
            catch (Exception)
            {
                // Fallback: Verwende das Standardsymbol, wenn es ein Problem gibt
                return SystemIcons.Application;
            }
        }

        // Verbesserte ClipboardNotification-Klasse mit Vorkehrungen gegen doppelte Events
        public static class ClipboardNotification
        {
            public static event EventHandler ClipboardChanged;

            private static ClipboardWatcher watcher;

            static ClipboardNotification()
            {
                watcher = new ClipboardWatcher();
                watcher.ClipboardChanged += (o, e) => OnClipboardChanged();
            }

            private static void OnClipboardChanged()
            {
                ClipboardChanged?.Invoke(null, EventArgs.Empty);
            }

            private class ClipboardWatcher : Form
            {
                public event EventHandler ClipboardChanged;

                // Debounce-Mechanismus auf Ebene des Watchers
                private DateTime lastEventTime = DateTime.MinValue;
                private readonly TimeSpan messageDebounceTime = TimeSpan.FromMilliseconds(100);

                public ClipboardWatcher()
                {
                    // Notwendig, damit die Form Nachrichten erhält, aber unsichtbar bleibt
                    ShowInTaskbar = false;
                    FormBorderStyle = FormBorderStyle.None;
                    Size = new Size(0, 0);
                    WindowState = FormWindowState.Minimized;

                    // Melde sich für Windows Clipboard Nachrichten an
                    SetClipboardViewer();
                }

                private void SetClipboardViewer()
                {
                    // Füge die Form zur Clipboard Chain hinzu
                    AddClipboardFormatListener(Handle);
                }

                protected override void WndProc(ref Message m)
                {
                    const int WM_CLIPBOARDUPDATE = 0x031D;

                    if (m.Msg == WM_CLIPBOARDUPDATE)
                    {
                        DateTime now = DateTime.Now;

                        // Prüfe, ob das Event innerhalb der Debounce-Zeit liegt
                        if ((now - lastEventTime) > messageDebounceTime)
                        {
                            lastEventTime = now;
                            ClipboardChanged?.Invoke(this, EventArgs.Empty);
                        }
                    }

                    base.WndProc(ref m);
                }

                [DllImport("user32.dll", SetLastError = true)]
                [return: MarshalAs(UnmanagedType.Bool)]
                private static extern bool AddClipboardFormatListener(IntPtr hwnd);
            }
        }
    }
}