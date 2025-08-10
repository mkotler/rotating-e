using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace RotatingEOverlay;

public class OutlookWatcher : IDisposable
{
    private dynamic? _app;           // Outlook.Application
    private dynamic? _ns;            // Namespace
    private dynamic? _inbox;         // MAPIFolder
    private Timer? _pollTimer;
    private string? _lastTopId;
    private bool _initialized;

    public event EventHandler? NewMail;

    public bool Initialize()
    {
        try
        {
            Type? t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null) { System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Outlook COM ProgID not found"); return false; }
            _app = Activator.CreateInstance(t);
            _ns = _app?.GetNamespace("MAPI");
            _ns?.Logon(); // default profile
            // 6 = olFolderInbox
            _inbox = _ns?.GetDefaultFolder(6);
            // Prime last id
            _lastTopId = GetTopMessageId();
            _pollTimer = new Timer(async _ => await PollAsync(), null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(5));
            _initialized = true;
            System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Initialized. First top id=" + (_lastTopId ?? "<none>"));
            return true;
        }
        catch
        {
            System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Initialize failed");
            return false;
        }
    }

    private async Task PollAsync()
    {
        try
        {
            string? currentTop = GetTopMessageId();
            if (currentTop != null && _lastTopId != null && currentTop != _lastTopId)
            {
                _lastTopId = currentTop;
                System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Detected new mail. New top id=" + currentTop);
                NewMail?.Invoke(this, EventArgs.Empty);
            }
            else if (_lastTopId == null && currentTop != null)
            {
                _lastTopId = currentTop;
                System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Primed top id=" + currentTop);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Poll: no change top=" + (currentTop ?? "<null>") );
            }
        }
        catch(Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("[OutlookWatcher] Poll exception: " + ex.Message);
        }
        await Task.CompletedTask;
    }

    private string? GetTopMessageId()
    {
        try
        {
            if (_inbox == null) return null;
            dynamic items = _inbox.Items;
            items.Sort("[ReceivedTime]", true); // descending
            dynamic? first = items.Count > 0 ? items[1] : null; // 1-based
            if (first == null) return null;
            string id = first.EntryID as string ?? string.Empty;
            return string.IsNullOrEmpty(id) ? null : id;
        }
        catch { return null; }
    }

    public void Dispose()
    {
        _pollTimer?.Dispose();
        ReleaseCom(_inbox);
        ReleaseCom(_ns);
        ReleaseCom(_app);
    }

    private static void ReleaseCom(object? o)
    {
        if (o != null && Marshal.IsComObject(o))
        {
            try { Marshal.ReleaseComObject(o); } catch { }
        }
    }
}
