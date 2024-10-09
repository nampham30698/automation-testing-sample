using AutomationTestingSample.Core.Helpers;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;

namespace AutomationTestingSample.Core.Reports;

public sealed class ExtentReporting
{
    private static ExtentReporting _instance = null;
    private static readonly object _myLock = new();

    private List<TestGroup> _testGroups = [];

    private TestGroup _currentTestGroup;


    private string _keyNode;

    private ExtentReports _extentReports;

    private ExtentReporting() { }

    public static ExtentReporting Instance
    {
        get
        {
            lock (_myLock)
            {
                if (_instance == null)
                {
                    _instance = new ExtentReporting();
                }
                return _instance;
            }
        }
    }

    /// <summary>
    /// Create ExtentReporting and attach ExtentHtmlReporter
    /// </summary>
    /// <returns></returns>
    private ExtentReports StartReporting()
    {
        var path = FileHelpers.ProjectPath + @"Results\";

        if (_extentReports == null)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            
            _extentReports = new ExtentReports();

            var extentSparkReporter = new ExtentSparkReporter(path + @"reports.html");

            extentSparkReporter.Config.DocumentTitle = "Automation Test Report";
            extentSparkReporter.Config.ReportName = "Automation Test Report";
            extentSparkReporter.Config.Encoding = "UTF-8";
            extentSparkReporter.Config.Theme = AventStack.ExtentReports.Reporter.Config.Theme.Standard;
            extentSparkReporter.Config.TimeStampFormat = "hh:mm:ss";

            _extentReports.AttachReporter(extentSparkReporter);
        }

        return _extentReports;
    }

    /// <summary>
    /// Create a new Test in reporter
    /// </summary>
    /// <param name="testNameGroup"></param>
    public void CreateTest(string testGroupName, string keyNode, string nodeName)
    {
        _keyNode = keyNode;

        _currentTestGroup = _testGroups.FirstOrDefault(x => x.Name.Equals(testGroupName, StringComparison.OrdinalIgnoreCase));

        if (_currentTestGroup is null)
        {
            var group = StartReporting().CreateTest(testGroupName);

            _currentTestGroup = new TestGroup()
            {
                Name = testGroupName,
                Group = group,
                Nodes = new Dictionary<string, ExtentTest>
                {
                    { keyNode, group.CreateNode(nodeName) }
                }
            };

            _testGroups.Add(_currentTestGroup);
        }
        else
        {
            _currentTestGroup.Nodes.Add(keyNode, _currentTestGroup.Group.CreateNode(nodeName));
        }
    }

    /// <summary>
    /// update/flush the info to reporter
    /// </summary>
    public void EndReporting()
    {
        StartReporting().Flush();
    }

    /// <summary>
    /// Log info message in report
    /// </summary>
    /// <param name="info"></param>
    public void LogInfo(string info)
    {
        _currentTestGroup.Nodes[_keyNode].Info(info);
    }

    /// <summary>
    /// Log pass message in report
    /// </summary>
    /// <param name="info"></param>
    public void LogPass(string info)
    {
        _currentTestGroup.Nodes[_keyNode].Pass(info);
    }

    /// <summary>
    /// Log fail message in report
    /// </summary>
    /// <param name="info"></param>
    public void LogFail(string info)
    {
        _currentTestGroup.Nodes[_keyNode].Fail(info);
    }

    public void LogFail(string info, string image)
    {
        _currentTestGroup.Nodes[_keyNode].Fail(info, MediaEntityBuilder.CreateScreenCaptureFromPath(image).Build());
    }

    public void LogFail(Exception ex)
    {
        _currentTestGroup.Nodes[_keyNode].Fail(ex);
    }

    public void LogFail(Exception ex, string image)
    {
        _currentTestGroup.Nodes[_keyNode].Fail(ex, MediaEntityBuilder.CreateScreenCaptureFromPath(image).Build());
    }

    /// <summary>
    /// Log screenshot in report
    /// </summary>
    /// <param name="info"></param>
    /// <param name="image"></param>
    public void LogScreenshot(string info, string image)
    {
        if (string.IsNullOrEmpty(image))
        {
            _currentTestGroup.Nodes[_keyNode].Info("No screenshot available.");
            return;
        }

        _currentTestGroup.Nodes[_keyNode].Info(info, MediaEntityBuilder.CreateScreenCaptureFromPath(image).Build());
    }

    public class TestGroup
    {
        public string Name { get; set; }

        public ExtentTest Group { get; set; }

        public Dictionary<string,ExtentTest> Nodes { get; set; }
    }
}