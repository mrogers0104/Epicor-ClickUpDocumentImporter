using ClickUpDocumentImporter.Helpers;
using System.Runtime.InteropServices;
using System.Text;
//using File = System.IO.File;

namespace ClickUpDocumentImporter.Helpers
{
    public enum LogLevel
    {
        Information,
        Success,
        Warning,
        Error,
        Debug
    }

    public class SelectionItem
    {
        public SelectionItem() { }

        public SelectionItem(string text, object value = null)
        {
            Text = text;
            Value = value ?? text;
            IsSelected = false;
        }

        public string Text { get; set; }
        public object Value { get; set; }
        public bool IsSelected { get; set; }
    }

    /// <summary> A comprehensive console utility class for your .NET 8 application with all the
    /// requested features plus some extras. Here's what's included: </summary> <remarks>
    ///
    /// Key Features:
    /// 1. <b>Logging with Colors & File Output</b>
    ///
    /// * Five log levels: Information(Cyan), Success(Green), Warning(Yellow), Error(Red), Debug(Gray)
    /// * Thread-safe logging to C:\temp\console_log.txt
    /// * Timestamps and log levels in file output
    ///
    /// 2. Selection Lists
    ///
    /// * Single Select: Arrow keys to navigate, Enter to select, Esc to cancel
    /// * Multi-Select: Space to toggle, arrow keys to navigate, checkboxes show selection state
    /// * Visual indicators with colors(cyan for current, green for selected)
    /// * Customizable indentation
    ///
    /// 3. User Input
    ///
    /// * AskQuestion() : String input with optional default values
    /// * AskYesNo(): Boolean input with Y/N prompt and default value
    ///
    /// 4. Modern Progress Bar
    ///
    /// * Filled portion in green(█), empty in dark gray(░)
    /// * Shows percentage, current/total count, and custom labels
    /// * RunWithProgress() wrapper for easy integration
    ///
    /// 5. Thread Safety
    ///
    /// * All methods use a lock object to prevent conflicts
    /// * Safe for concurrent operations from multiple threads
    ///
    /// 6. Error Handling
    ///
    /// * Try-catch blocks in all methods
    /// * Errors logged to console and file
    /// * Graceful degradation if log file unavailable
    ///
    /// Bonus Features:
    ///
    /// * WriteHeader() : Styled section headers
    /// * WriteSeparator(): Visual separators
    /// * Clear() : Thread-safe console clearing
    /// * Pause(): "Press any key to continue" functionality
    ///
    /// Usage Tips:
    ///
    /// * The log file is automatically created at C:\temp\console_log.txt
    /// * All user interactions are logged for audit purposes
    /// * The example program demonstrates all features comprehensively
    /// * Selection lists automatically handle edge cases(empty lists, navigation wrapping)
    /// </remarks>
    public static class ConsoleHelper
    {
        private static readonly object _lock = new object();
        private static readonly string _logDirectory = @"C:\temp\ClickUpDocumentInport";
        private static string _logFilename = string.Empty;
        private static string _logFilePath = @"C:\temp\console_log.txt";
        private static bool _logFileInitialized = false;

        // Color mappings for log levels
        private static readonly Dictionary<LogLevel, ConsoleColor> _logColors = new()
        {
            { LogLevel.Information, ConsoleColor.Cyan },
            { LogLevel.Success, ConsoleColor.Green },
            { LogLevel.Warning, ConsoleColor.Yellow },
            { LogLevel.Error, ConsoleColor.Red },
            { LogLevel.Debug, ConsoleColor.Gray }
        };

        static ConsoleHelper()
        {
            InitializeLogFile();
        }

        #region Logging Methods

        private static void InitializeLogFile()
        {
            try
            {
                _logFilename = $"console_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                _logFilePath = Path.Combine(_logDirectory, _logFilename);
                //string directory = Path.GetDirectoryName(_logFilePath);
                string directory = _logDirectory;
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (!File.Exists(_logFilePath))
                {
                    File.WriteAllText(_logFilePath, $"=== Console Log Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={Environment.NewLine}");
                }

                _logFileInitialized = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not initialize log file: {ex.Message}");
                _logFileInitialized = false;
            }
        }

        public static void Log(string message, LogLevel level = LogLevel.Information, bool writeToConsole = false)
        {
            lock (_lock)
            {
                try
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string formattedMessage = $"[{timestamp}] [{level,-11}] {message}";

                    // Console output with color (only if requested)
                    if (writeToConsole)
                    {
                        ConsoleColor originalColor = Console.ForegroundColor;
                        Console.ForegroundColor = _logColors[level];
                        Console.WriteLine(formattedMessage);
                        Console.ForegroundColor = originalColor;
                    }

                    // File output
                    if (_logFileInitialized)
                    {
                        File.AppendAllText(_logFilePath, formattedMessage + Environment.NewLine);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error writing to log: {ex.Message}");
                }
            }
        }

        public static void LogInformation(string message, bool writeToConsole = false) => Log(message, LogLevel.Information, writeToConsole);

        public static void LogSuccess(string message, bool writeToConsole = false) => Log(message, LogLevel.Success, writeToConsole);

        public static void LogWarning(string message, bool writeToConsole = false) => Log(message, LogLevel.Warning, writeToConsole);

        public static void LogError(string message, bool writeToConsole = false) => Log(message, LogLevel.Error, writeToConsole);

        public static void LogDebug(string message, bool writeToConsole = false) => Log(message, LogLevel.Debug, writeToConsole);

        #endregion Logging Methods

        #region Selection Methods

        public static SelectionItem? SelectFromList(List<SelectionItem> items, string prompt = "Select an item:", int indent = 2)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("Items list cannot be null or empty");
            }

            lock (_lock)
            {
                try
                {
                    int promptTop = Console.CursorTop;

                    //Console.WriteLine(prompt);
                    ConsoleHelper.WriteInfo(prompt);
                    int selectedIndex = 0;
                    bool selecting = true;
                    int startTop = Console.CursorTop;

                    List<int> pos = new List<int>();

                    // Draw initial list
                    for (int i = 0; i < items.Count; i++)
                    {
                        Console.WriteLine();
                    }

                    var now = Console.CursorTop;


                    while (selecting)
                    {
                        DisplaySelectionList(items, selectedIndex, indent, false, startTop);
                        var p = Console.CursorTop;
                        pos.Add(p);

                        ConsoleKeyInfo key = Console.ReadKey(true);
                        switch (key.Key)
                        {
                            case ConsoleKey.UpArrow:
                                selectedIndex = selectedIndex > 0 ? selectedIndex - 1 : items.Count - 1;
                                break;
                            case ConsoleKey.DownArrow:
                                selectedIndex = selectedIndex < items.Count - 1 ? selectedIndex + 1 : 0;
                                break;
                            case ConsoleKey.Enter:
                                selecting = false;
                                break;
                            case ConsoleKey.Escape:
                                ClearSelectionDisplay(items.Count + 1, promptTop);
                                return null;
                        }
                    }

                    now = Console.CursorTop;

                    int newTop = promptTop - selectedIndex - 2;
                    newTop = newTop < 0 ? 0 : newTop;

                    ClearSelectionDisplay(items.Count + 3, newTop);
                    //ClearSelectionDisplay(items.Count + 1, promptTop);
                    var selected = items[selectedIndex];
                    var selectedText = (string) (selected.Value);
                    //Console.WriteLine($"{new string(' ', indent)}✓ Selected: {selected.Text}");
                    ConsoleHelper.WriteInfo($"✓ Selected: {selectedText.Trim()}");
                    //Log($"User selected: {selected.Text}", LogLevel.Information);
                    return selected;
                }
                catch (Exception ex)
                {
                    LogError($"Error in SelectFromList: {ex.Message}");
                    throw;
                }
            }
        }

        public static List<SelectionItem> MultiSelectFromList(List<SelectionItem> items, string prompt = "Select items (Space to toggle, Enter to confirm):", int indent = 2)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("Items list cannot be null or empty");
            }

            lock (_lock)
            {
                try
                {
                    int promptTop = Console.CursorTop;

                    Console.WriteLine(prompt);
                    int currentIndex = 0;
                    bool selecting = true;
                    int startTop = Console.CursorTop;

                    // Draw initial list
                    for (int i = 0; i < items.Count; i++)
                    {
                        Console.WriteLine();
                    }

                    while (selecting)
                    {
                        DisplaySelectionList(items, currentIndex, indent, true, startTop);

                        ConsoleKeyInfo key = Console.ReadKey(true);
                        switch (key.Key)
                        {
                            case ConsoleKey.UpArrow:
                                currentIndex = currentIndex > 0 ? currentIndex - 1 : items.Count - 1;
                                break;
                            case ConsoleKey.DownArrow:
                                currentIndex = currentIndex < items.Count - 1 ? currentIndex + 1 : 0;
                                break;
                            case ConsoleKey.Spacebar:
                                items[currentIndex].IsSelected = !items[currentIndex].IsSelected;
                                break;
                            case ConsoleKey.Enter:
                                selecting = false;
                                break;
                            case ConsoleKey.Escape:
                                ClearSelectionDisplay(items.Count + 1, promptTop);
                                return new List<SelectionItem>();
                        }
                    }

                    ClearSelectionDisplay(items.Count + 1, promptTop);

                    var selected = items.Where(i => i.IsSelected).ToList();
                    Console.WriteLine($"{new string(' ', indent)}X Selected {selected.Count} item(s)");
                    Log($"User selected {selected.Count} items: {string.Join(", ", selected.Select(s => s.Text))}", LogLevel.Information);
                    return selected;
                }
                catch (Exception ex)
                {
                    LogError($"Error in MultiSelectFromList: {ex.Message}");
                    throw;
                }
            }
        }

        private static void DisplaySelectionList(List<SelectionItem> items, int selectedIndex, int indent, bool multiSelect, int startTop)
        {
            Console.SetCursorPosition(0, Console.CursorTop - items.Count);
            //Console.SetCursorPosition(0, startTop);

            for (int i = 0; i < items.Count; i++)
            {
                //// Clear the entire line first
                //Console.Write(new string(' ', Console.WindowWidth - 1));
                //Console.SetCursorPosition(0, startTop + i);

                string indentation = new string(' ', indent);
                string selector = i == selectedIndex ? ">" : " ";
                string checkbox = multiSelect ? (items[i].IsSelected ? "[✓]" : "[ ]") : "";

                if (i == selectedIndex)
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.Write($"{indentation}{selector} ");
                    if (multiSelect) Console.Write($"{checkbox} ");
                    Console.WriteLine($"{items[i].Text}".PadRight(Console.WindowWidth - indent - 5));
                    Console.ResetColor();
                }
                else
                {
                    Console.Write($"{indentation}{selector} ");
                    if (multiSelect)
                    {
                        if (items[i].IsSelected)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{checkbox} ");
                            Console.ResetColor();
                        }
                        else
                        {
                            Console.Write($"{checkbox} ");
                        }
                    }
                    Console.WriteLine($"{items[i].Text}".PadRight(Console.WindowWidth - indent - 5));
                }
            }
        }

        private static void ClearSelectionDisplay(int lines, int startTop)
        {
            Console.SetCursorPosition(0, startTop);
            for (int i = 0; i < lines; i++)
            {
                Console.Write(new string(' ', Console.WindowWidth - 1));
                Console.WriteLine();
            }
            Console.SetCursorPosition(0, startTop);
        }

        #endregion Selection Methods

        #region Input Methods

        public static string AskQuestion(string question, string defaultValue = null)
        {
            lock (_lock)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    if (string.IsNullOrEmpty(defaultValue))
                    {
                        Console.Write($"? {question}: ");
                    }
                    else
                    {
                        Console.Write($"? {question} [{defaultValue}]: ");
                    }
                    Console.ResetColor();

                    string input = Console.ReadLine();
                    string result = string.IsNullOrWhiteSpace(input) ? defaultValue : input;

                    Log($"Question asked: '{question}' - Answer: '{result}'", LogLevel.Debug);
                    return result;
                }
                catch (Exception ex)
                {
                    LogError($"Error in AskQuestion: {ex.Message}");
                    throw;
                }
            }
        }

        public static bool AskYesNo(string question, bool defaultValue = true)
        {
            lock (_lock)
            {
                try
                {
                    string defaultIndicator = defaultValue ? "Y/n" : "y/N";
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write($"? {question} ({defaultIndicator}): ");
                    Console.ResetColor();

                    string input = Console.ReadLine()?.Trim().ToLower();

                    bool result;
                    if (string.IsNullOrEmpty(input))
                    {
                        result = defaultValue;
                    }
                    else if (input == "y" || input == "yes")
                    {
                        result = true;
                    }
                    else if (input == "n" || input == "no")
                    {
                        result = false;
                    }
                    else
                    {
                        result = defaultValue;
                    }

                    Log($"Yes/No question: '{question}' - Answer: {result}", LogLevel.Debug);
                    return result;
                }
                catch (Exception ex)
                {
                    LogError($"Error in AskYesNo: {ex.Message}");
                    throw;
                }
            }
        }

        #endregion Input Methods

        #region Progress Bar

        public static void ShowProgressBar(int current, int total, string label = "", int barWidth = 40)
        {
            lock (_lock)
            {
                try
                {
                    if (total <= 0) return;

                    double percentage = (double)current / total;
                    int filledWidth = (int)(barWidth * percentage);
                    int emptyWidth = barWidth - filledWidth;

                    StringBuilder bar = new StringBuilder();
                    bar.Append("[");

                    // Filled portion
                    Console.Write("[");
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(new string('█', filledWidth));
                    Console.ResetColor();

                    // Empty portion
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.Write(new string('░', emptyWidth));
                    Console.ResetColor();
                    Console.Write("]");

                    // Percentage and label
                    Console.Write($" {percentage:P0} ");

                    if (!string.IsNullOrEmpty(label))
                    {
                        Console.Write($"- {label}");
                    }

                    Console.Write($" ({current}/{total})");

                    // Move cursor back to beginning of line for next update
                    if (current < total)
                    {
                        Console.Write("\r");
                    }
                    else
                    {
                        Console.WriteLine(); // Complete
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Error in ShowProgressBar: {ex.Message}");
                }
            }
        }

        public static void RunWithProgress(Action<Action<int, string>> action, int totalSteps, string initialLabel = "Processing...")
        {
            try
            {
                Log($"Started progress task: {initialLabel}", LogLevel.Information);

                action((progress, label) =>
                {
                    ShowProgressBar(progress, totalSteps, label);
                });

                Log($"Completed progress task: {initialLabel}", LogLevel.Success);
            }
            catch (Exception ex)
            {
                LogError($"Error in RunWithProgress: {ex.Message}");
                throw;
            }
        }

        #endregion Progress Bar

        #region Utility Methods

        public static void WriteHeader(string title)
        {
            lock (_lock)
            {
                try
                {
                    Console.WriteLine(); // Ensure we start on a new line
                    int width = Math.Min(Console.WindowWidth - 1, 80);
                    string border = new string('═', width);

                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine(border);
                    Console.WriteLine(title.PadLeft((width + title.Length) / 2).PadRight(width));
                    Console.WriteLine(border);
                    Console.ResetColor();

                    Log($"Header displayed: {title}", LogLevel.Debug);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteHeader: {ex.Message}");
                }
            }
        }

        public static void Write(string msg, LogLevel logLevel)
        {
            lock (_lock)
            {
                try
                {
                    //Console.WriteLine();
                    Console.ForegroundColor = _logColors[logLevel];
                    Console.WriteLine(msg);
                    Console.ResetColor();

                    Log($"{logLevel} displayed: {msg}", LogLevel.Debug);
                }
                catch (Exception ex)
                {
                    LogError($"Error in Write: {ex.Message}");
                }
            }
        }

        public static void WriteInfo(string msg)
        {
            lock (_lock)
            {
                try
                {
                    Write(msg, LogLevel.Information);
                    LogInformation(msg);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteInfo: {ex.Message}");
                }
            }
        }

        public static void WriteSuccess(string msg)
        {
            lock (_lock)
            {
                try
                {
                    Write(msg, LogLevel.Success);
                    LogSuccess(msg);

                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteInfo: {ex.Message}");
                }
            }
        }

        public static void WriteWarning(string msg)
        {
            lock (_lock)
            {
                try
                {
                    Console.Write(msg, LogLevel.Warning);
                    LogWarning(msg);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteInfo: {ex.Message}");
                }
            }
        }
        public static void WriteError(string msg)
        {
            lock (_lock)
            {
                try
                {
                    Write(msg, LogLevel.Error);
                    LogError(msg);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteInfo: {ex.Message}");
                }
            }
        }

        public static void WriteBlank()
        {
            lock (_lock)
            {
                try
                {
                    string msg = string.Empty;
                    Write(msg, LogLevel.Information);
                    LogError(msg);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteBlank: {ex.Message}");
                }
            }
        }

        public static void WriteLogPath()
        {
            lock (_lock)
            {
                try
                {
                    string msg = $"Log file located at: {_logFilePath}";
                    Write(msg, LogLevel.Success);
                    LogError(msg);
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteLogPath: {ex.Message}");
                }
            }
        }


        public static void WriteSeparator(char character = '~', int length = 0)
        {
            lock (_lock)
            {
                try
                {
                    int lineLength = length > 0 ? length : Console.WindowWidth - 6;
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine(new string(character, lineLength));
                    Console.ResetColor();
                }
                catch (Exception ex)
                {
                    LogError($"Error in WriteSeparator: {ex.Message}");
                }
            }
        }

        public static void Clear()
        {
            lock (_lock)
            {
                Console.Clear();
                Log("Console cleared", LogLevel.Debug);
            }
        }

        public static void Pause(string message = "Press any key to continue...")
        {
            lock (_lock)
            {
                try
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    Console.WriteLine();
                    Console.Write(message);
                    Console.ResetColor();
                    Console.ReadKey(true);
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    LogError($"Error in Pause: {ex.Message}");
                }
            }
        }

        #endregion Utility Methods
    }
}

// *** Example Program Below ***

//using System;
//using System.Collections.Generic;
//using System.Threading;
//using System.Threading.Tasks;
//using ConsoleUtilities;

//namespace ConsoleDemo
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            try
//            {
//                // Display header
//                ConsoleHelper.WriteHeader("Console Helper Demo Application");
//                ConsoleHelper.Log("Application started", LogLevel.Information);

//                // Demo 1: Logging with different levels
//                DemoLogging();
//                ConsoleHelper.Pause();

//                // Demo 2: Single selection
//                ConsoleHelper.Clear();
//                DemoSingleSelect();
//                ConsoleHelper.Pause();

//                // Demo 3: Multi-selection
//                ConsoleHelper.Clear();
//                DemoMultiSelect();
//                ConsoleHelper.Pause();

//                // Demo 4: Question and Yes/No
//                ConsoleHelper.Clear();
//                DemoQuestions();
//                ConsoleHelper.Pause();

//                // Demo 5: Progress bar
//                ConsoleHelper.Clear();
//                DemoProgressBar();
//                ConsoleHelper.Pause();

//                // Demo 6: Thread safety
//                ConsoleHelper.Clear();
//                DemoThreadSafety();
//                ConsoleHelper.Pause();

//                ConsoleHelper.WriteHeader("Demo Complete");
//                ConsoleHelper.LogSuccess("All demos completed successfully!");
//                ConsoleHelper.LogInformation($"Check the log file at: C:\\temp\\console_log.txt");
//            }
//            catch (Exception ex)
//            {
//                ConsoleHelper.LogError($"Application error: {ex.Message}");
//                ConsoleHelper.LogDebug($"Stack trace: {ex.StackTrace}");
//            }
//        }

//        static void DemoLogging()
//        {
//            ConsoleHelper.WriteHeader("Demo 1: Logging with Color");
//            ConsoleHelper.WriteSeparator();

//            ConsoleHelper.LogInformation("This is an informational message");
//            ConsoleHelper.LogSuccess("This is a success message");
//            ConsoleHelper.LogWarning("This is a warning message");
//            ConsoleHelper.LogError("This is an error message");
//            ConsoleHelper.LogDebug("This is a debug message");

//            ConsoleHelper.WriteSeparator();
//            ConsoleHelper.LogInformation("All messages are also logged to C:\\temp\\console_log.txt");
//        }

//        static void DemoSingleSelect()
//        {
//            ConsoleHelper.WriteHeader("Demo 2: Single Selection");
//            ConsoleHelper.WriteSeparator();

//            var items = new List<SelectionItem>
//            {
//                new SelectionItem("Option 1: Create new project", "create"),
//                new SelectionItem("Option 2: Open existing project", "open"),
//                new SelectionItem("Option 3: Configure settings", "settings"),
//                new SelectionItem("Option 4: View documentation", "docs"),
//                new SelectionItem("Option 5: Exit application", "exit")
//            };

//            ConsoleHelper.LogInformation("Use Arrow Keys to navigate, Enter to select, Esc to cancel");
//            ConsoleHelper.WriteSeparator();

//            var selected = ConsoleHelper.SelectFromList(items, "Please select an action:", 2);

//            if (selected != null)
//            {
//                ConsoleHelper.WriteSeparator();
//                ConsoleHelper.LogSuccess($"You selected: {selected.Text} (Value: {selected.Value})");
//            }
//            else
//            {
//                ConsoleHelper.LogWarning("Selection was cancelled");
//            }
//        }

//        static void DemoMultiSelect()
//        {
//            ConsoleHelper.WriteHeader("Demo 3: Multi-Selection");
//            ConsoleHelper.WriteSeparator();

//            var features = new List<SelectionItem>
//            {
//                new SelectionItem("Authentication Module"),
//                new SelectionItem("Database Integration"),
//                new SelectionItem("API Gateway"),
//                new SelectionItem("Caching Layer"),
//                new SelectionItem("Logging Service"),
//                new SelectionItem("Message Queue")
//            };

//            ConsoleHelper.LogInformation("Use Arrow Keys to navigate, Space to toggle, Enter to confirm, Esc to cancel");
//            ConsoleHelper.WriteSeparator();

//            var selected = ConsoleHelper.MultiSelectFromList(
//                features,
//                "Select the features you want to install:",
//                2
//            );

//            ConsoleHelper.WriteSeparator();
//            if (selected.Count > 0)
//            {
//                ConsoleHelper.LogSuccess($"Selected {selected.Count} feature(s):");
//                foreach (var item in selected)
//                {
//                    ConsoleHelper.LogInformation($"  • {item.Text}");
//                }
//            }
//            else
//            {
//                ConsoleHelper.LogWarning("No features selected");
//            }
//        }

//        static void DemoQuestions()
//        {
//            ConsoleHelper.WriteHeader("Demo 4: User Input");
//            ConsoleHelper.WriteSeparator();

//            // String input
//            string name = ConsoleHelper.AskQuestion("What is your name?", "Anonymous");
//            ConsoleHelper.LogInformation($"Hello, {name}!");

//            // String input with default
//            string projectName = ConsoleHelper.AskQuestion("Enter project name", "MyProject");
//            ConsoleHelper.LogInformation($"Project name set to: {projectName}");

//            ConsoleHelper.WriteSeparator();

//            // Yes/No questions
//            bool enableDebug = ConsoleHelper.AskYesNo("Enable debug mode?", false);
//            ConsoleHelper.LogInformation($"Debug mode: {(enableDebug ? "Enabled" : "Disabled")}");

//            bool confirmAction = ConsoleHelper.AskYesNo("Do you want to continue?", true);

//            if (confirmAction)
//            {
//                ConsoleHelper.LogSuccess("User confirmed action");
//            }
//            else
//            {
//                ConsoleHelper.LogWarning("User cancelled action");
//            }
//        }

//        static void DemoProgressBar()
//        {
//            ConsoleHelper.WriteHeader("Demo 5: Progress Bar");
//            ConsoleHelper.WriteSeparator();

//            // Simple progress bar
//            ConsoleHelper.LogInformation("Processing files...");
//            for (int i = 0; i <= 100; i += 5)
//            {
//                ConsoleHelper.ShowProgressBar(i, 100, $"File {i}/100");
//                Thread.Sleep(100); // Simulate work
//            }

//            ConsoleHelper.WriteSeparator();

//            // Progress bar with action wrapper
//            ConsoleHelper.LogInformation("Running complex task...");
//            ConsoleHelper.RunWithProgress((updateProgress) =>
//            {
//                for (int i = 1; i <= 50; i++)
//                {
//                    // Simulate different stages
//                    string stage = i <= 10 ? "Initializing" :
//                                   i <= 30 ? "Processing" :
//                                   i <= 45 ? "Finalizing" : "Completing";

//                    updateProgress(i, stage);
//                    Thread.Sleep(50);
//                }
//            }, 50, "Complex Operation");

//            ConsoleHelper.LogSuccess("All tasks completed!");
//        }

//        static void DemoThreadSafety()
//        {
//            ConsoleHelper.WriteHeader("Demo 6: Thread Safety");
//            ConsoleHelper.WriteSeparator();

//            ConsoleHelper.LogInformation("Starting 5 concurrent threads...");
//            ConsoleHelper.WriteSeparator();

//            var tasks = new List<Task>();

//            for (int i = 1; i <= 5; i++)
//            {
//                int threadId = i;
//                tasks.Add(Task.Run(() =>
//                {
//                    for (int j = 1; j <= 5; j++)
//                    {
//                        ConsoleHelper.Log($"Thread {threadId}: Message {j}/5",
//                            j % 2 == 0 ? LogLevel.Information : LogLevel.Success);
//                        Thread.Sleep(Random.Shared.Next(50, 150));
//                    }
//                }));
//            }

//            Task.WaitAll(tasks.ToArray());

//            ConsoleHelper.WriteSeparator();
//            ConsoleHelper.LogSuccess("All threads completed without conflicts!");
//            ConsoleHelper.LogInformation("Notice how messages are properly synchronized");
//        }
//    }
//}
