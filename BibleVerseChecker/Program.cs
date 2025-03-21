using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace BibleVerseOpener
{
    class Program
    {
        // Importiere die benötigten Win32 API Funktionen
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        // Definiere Konstanten
        private const int MOD_CONTROL = 0x0002;
        private const int VK_C = 0x43;
        private const int HOTKEY_ID = 1;

        // Bibelreferenz-Regex
        private static readonly Regex BibleVerseRegex = new Regex(
            @"(1\s*|2\s*|3\s*)?(Mose|Samuel|Könige|Chronik|Johannes|Timotheus|Petrus|Thessalonicher|Korinther)?\.?\s*" +
            @"(Genesis|Gen|Exodus|Ex|Levitikus|Lev|Numeri|Num|Deuteronomium|Dtn|Deutero|Josua|Jos|Richter|Ri|Rut|Rt|" +
            @"Samuel|Sam|Könige|Kön|Chronik|Chr|Esra|Esr|Nehemia|Neh|Tobit|Tob|Judit|Jdt|Ester|Est|" +
            @"Makkabäer|Makk|Ijob|Hiob|Psalmen|Ps|Psalm|Sprüche|Spr|Kohelet|Koh|Hoheslied|Hld|Weisheit|Weish|" +
            @"Jesus Sirach|Sir|Jesaja|Jes|Jeremia|Jer|Klagelieder|Klgl|Baruch|Bar|Ezechiel|Ez|Daniel|Dan|" +
            @"Hosea|Hos|Joel|Am|Obadja|Obd|Jona|Jon|Micha|Mi|Nahum|Nah|Habakuk|Hab|Zefanja|Zef|Haggai|Hag|" +
            @"Sacharja|Sach|Maleachi|Mal|Matthäus|Mt|Markus|Mk|Lukas|Lk|Johannes|Joh|Apostelgeschichte|Apg|" +
            @"Römer|Röm|Korinther|Kor|Galater|Gal|Epheser|Eph|Philipper|Phil|Kolosser|Kol|Thessalonicher|Thess|" +
            @"Timotheus|Tim|Titus|Tit|Philemon|Phlm|Hebräer|Hebr|Jakobus|Jak|Petrus|Petr|Judas|Jud|Offenbarung|Offb)" +
            @"\s*(\d+)(?:\s*,\s*(\d+)(?:\s*-\s*(\d+))?)?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static DateTime lastCopyTime = DateTime.MinValue;
        private static NotifyIcon notifyIcon;

        [STAThread]
        static void Main(string[] args)
        {
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
            contextMenu.Items.Add("Info", null, (s, e) => ShowInfo());
            notifyIcon.ContextMenuStrip = contextMenu;

            // Doppelklick auf Icon zeigt Info an
            notifyIcon.DoubleClick += (s, e) => ShowInfo();

            // Registriere den Hotkey
            HookClipboard();

            // Zeige eine Benachrichtigung, dass die App gestartet wurde
            notifyIcon.ShowBalloonTip(3000, "Bibelvers-Opener", "Anwendung läuft im Hintergrund. Drücke 2x STRG+C innerhalb von 2 Sekunden, um einen Bibelvers zu öffnen.", ToolTipIcon.Info);

            // Halte die Anwendung am Laufen
            Application.Run();

            // Aufräumen beim Beenden
            notifyIcon.Visible = false;
            UnhookClipboard();
        }

        private static void ShowInfo()
        {
            MessageBox.Show(
                "Bibelvers-Opener\n\n" +
                "Markiere einen Bibelvers und drücke zweimal STRG+C innerhalb von 2 Sekunden.\n" +
                "Das Programm erkennt den Vers und öffnet ihn in deinem Browser.\n\n" +
                "Unterstützte Formate:\n" +
                "- Johannes 3,16\n" +
                "- Joh 3,16\n" +
                "- Joh 3,16-18\n" +
                "- 1. Korinther 13,4-7",
                "Über Bibelvers-Opener",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
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

        private static void OnClipboardChanged(object sender, EventArgs e)
        {
            try
            {
                // Prüfen, ob das zweite STRG+C innerhalb von 2 Sekunden erfolgte
                DateTime now = DateTime.Now;
                if ((now - lastCopyTime).TotalSeconds <= 2)
                {
                    // Hole den Text aus der Zwischenablage
                    string clipboardText = Clipboard.GetText();
                    if (!string.IsNullOrWhiteSpace(clipboardText))
                    {
                        ProcessBibleVerses(clipboardText);
                    }
                    lastCopyTime = DateTime.MinValue; // Zurücksetzen
                }
                else
                {
                    lastCopyTime = now;
                }
            }
            catch (Exception ex)
            {
                notifyIcon.ShowBalloonTip(3000, "Fehler", $"Fehler beim Verarbeiten der Zwischenablage: {ex.Message}", ToolTipIcon.Error);
            }
        }

        private static void ProcessBibleVerses(string text)
        {
            var matches = BibleVerseRegex.Matches(text);

            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    // Extrahiere die gefundenen Gruppen
                    string numberPrefix = match.Groups[1].Value.Trim();
                    string bookAlt = match.Groups[2].Value.Trim();
                    string book = match.Groups[3].Value.Trim();
                    string chapter = match.Groups[4].Value.Trim();
                    string verse = match.Groups[5].Value.Trim();
                    string endVerse = match.Groups[6].Value.Trim();

                    // Kombiniere Präfix und Buch
                    string fullBook = !string.IsNullOrEmpty(numberPrefix) ?
                        $"{numberPrefix} {(!string.IsNullOrEmpty(bookAlt) ? bookAlt : book)}" :
                        book;

                    // Baue den Link auf
                    string url = $"https://www.bibleserver.com/LUT/{Uri.EscapeDataString(fullBook)}{chapter}";

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
                    notifyIcon.ShowBalloonTip(2000, "Bibelvers gefunden",
                        $"Öffne {fullBook} {chapter}{(!string.IsNullOrEmpty(verse) ? $",{verse}" : "")}" +
                        $"{(!string.IsNullOrEmpty(endVerse) ? $"-{endVerse}" : "")}",
                        ToolTipIcon.Info);

                    // Nur den ersten gefundenen Vers verarbeiten
                    break;
                }
            }
            else
            {
                notifyIcon.ShowBalloonTip(2000, "Kein Bibelvers gefunden",
                    "Im kopierten Text wurde kein Bibelvers erkannt.", ToolTipIcon.Warning);
            }
        }
    }

    // Hilfsklasse zur Überwachung der Zwischenablage
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
                    ClipboardChanged?.Invoke(this, EventArgs.Empty);
                }

                base.WndProc(ref m);
            }

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool AddClipboardFormatListener(IntPtr hwnd);
        }
    }
}