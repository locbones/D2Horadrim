using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Reflection.Metadata.BlobBuilder;
using System.Data.Common;

namespace D2_Horadrim
{
    public class WorkspaceEntry
    {
        public string Name { get; set; } = "";
        public string FolderPath { get; set; } = "";
        public int? ColorArgb { get; set; }
    }

    public class WorkspaceTag
    {
        public string FolderPath { get; set; } = "";
        public Color BackColor { get; set; }
    }

    public class TabTag
    {
        public string FilePath { get; set; } = "";
        public string? WorkspaceFolderPath { get; set; }
        public Color? WorkspaceBackColor { get; set; }
    }

    public class WorkspacesConfig
    {
        public List<WorkspaceEntry> Workspaces { get; set; } = new();
    }

    public class PaintTabEventArgs : EventArgs
    {
        public required Graphics Graphics { get; init; }
        public Rectangle Bounds { get; init; }
        public bool Selected { get; init; }
        public required TabPage Page { get; init; }
        public int TabIndex { get; init; }
    }

    // TabControl with fully flat tabs: UserPaint so we draw the entire tab strip ourselves (no native 3D or borders).
    public class ThemedTabControl : TabControl
    {
        private const int WM_ERASEBKGND = 0x0014;
        private const int TCM_FIRST = 0x1300;
        private const int TCM_ADJUSTRECT = TCM_FIRST + 40;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public Color? TabStripBackColor { get; set; }
        public event EventHandler<PaintTabEventArgs>? PaintTab;

        public ThemedTabControl()
        {
            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.DoubleBuffer, true);
            SetStyle(ControlStyles.ResizeRedraw, true);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            if (TabStripBackColor.HasValue)
            {
                using (var brush = new SolidBrush(TabStripBackColor.Value))
                    e.Graphics.FillRectangle(brush, e.ClipRectangle);
                return;
            }
            base.OnPaintBackground(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            int headerHeight = TabCount > 0 ? GetTabRect(TabCount - 1).Bottom + 2 : 0;
            Color stripBack = TabStripBackColor ?? BackColor;
            if (stripBack == Color.Empty || stripBack.A == 0)
                stripBack = SystemColors.Control;
            using (var brush = new SolidBrush(stripBack))
                e.Graphics.FillRectangle(brush, 0, 0, Width, headerHeight);
            for (int i = 0; i < TabCount; i++)
            {
                Rectangle bounds = GetTabRect(i);
                TabPage page = TabPages[i];
                bool selected = (SelectedIndex == i);
                PaintTab?.Invoke(this, new PaintTabEventArgs { Graphics = e.Graphics, Bounds = bounds, Selected = selected, Page = page, TabIndex = i });
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == TCM_ADJUSTRECT)
            {
                RECT rc = Marshal.PtrToStructure<RECT>(m.LParam);
                rc.Left -= 4;
                rc.Right += 4;
                rc.Top -= 2;
                rc.Bottom += 4;
                Marshal.StructureToPtr(rc, m.LParam, false);
            }
            else if (m.Msg == WM_ERASEBKGND && TabStripBackColor.HasValue)
            {
                using (var g = Graphics.FromHdc(m.WParam))
                using (var brush = new SolidBrush(TabStripBackColor.Value))
                    g.FillRectangle(brush, ClientRectangle);
                m.Result = (IntPtr)1;
                return;
            }
            base.WndProc(ref m);
        }
    }

    public class EditHistoryItem
    {
        public string DisplayText { get; }
        public string? ColumnName { get; }
        public string? RowName { get; }
        public string? OldValue { get; }
        public string? NewValue { get; }
        public string? PasteValue { get; }
        public string? PasteCount { get; }
        public string? PasteCellRange { get; }
        public string? MathOperation { get; }
        public string? MathValue { get; }
        public string? MathColumnName { get; }
        public string? MathCellRange { get; }
        // For round operations: "up" or "down".
        public string? MathRoundDirection { get; }
        // Math format kind: AddSub, MulDiv, Round, Custom, IncrementFill.
        public string? MathFormatKind { get; }

        public EditHistoryItem(string displayText, string? columnName = null, string? rowName = null, string? oldValue = null, string? newValue = null, string? pasteValue = null, string? pasteCount = null, string? pasteCellRange = null, string? mathOperation = null, string? mathValue = null, string? mathColumnName = null, string? mathCellRange = null, string? mathRoundDirection = null, string? mathFormatKind = null)
        {
            DisplayText = displayText;
            ColumnName = columnName;
            RowName = rowName;
            OldValue = oldValue;
            NewValue = newValue;
            PasteValue = pasteValue;
            PasteCount = pasteCount;
            PasteCellRange = pasteCellRange;
            MathOperation = mathOperation;
            MathValue = mathValue;
            MathColumnName = mathColumnName;
            MathCellRange = mathCellRange;
            MathRoundDirection = mathRoundDirection;
            MathFormatKind = mathFormatKind;
        }
    }

    public partial class D2Horadrim : Form, IMessageFilter
    {
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;
        private const int VK_MENU = 0x12;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_APP = 0x8000;
        private const uint WM_HOTKEY_ALT = WM_APP + 1;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr _lowLevelHookHandle = IntPtr.Zero;
        private LowLevelKeyboardProc? _lowLevelHookProc; // keep alive for GC
        private bool _altKeyLogicalDown; // set when we consume Alt key down so Alt+Click works

        private enum MathOp { Add, Subtract, Divide, Multiply, RoundUp, RoundDown, CustomFormula, IncrementFill }

        private const string RegistryKeyPath = @"Software\D2_Horadrim";
        private const string OriginalIndexColumnName = "__OriginalIndex";
        // GitHub repo for update checks: "owner/repo".
        private const string GitHubReleasesRepo = "locbones/D2Horadrim";
        private const int ProgressStreamReaderBufferSize = 1024;
        private const int MaxWorkspacePanelWidth = 260;
        private Image lockImage = Properties.Resources.Locked;
        ThemedTabControl tabControl = new ThemedTabControl();
        private Dictionary<string, DataTable> fileCache = new Dictionary<string, DataTable>(StringComparer.OrdinalIgnoreCase);
        private DataGridView? dataGridView;
        private TreeView treeView;
        private string currentFilePath;
        private string lastOpenedDirectory;
        private string? _startupFileArgument;
        private string? _startupWorkspaceRole;
        private string? _startupWorkspaceFolder;
        private HashSet<int> lockedColumns = new HashSet<int>();
        private HashSet<int> lockedRows = new HashSet<int>();
        private int _anchorColumnIndex = -1;
        private int _anchorRowIndex = -1;
        internal static readonly Color[] DefaultWorkspaceColors = new[] { Color.FromArgb(205, 168, 255), Color.FromArgb(253, 176, 253), Color.FromArgb(253, 168, 181), Color.FromArgb(179, 253, 168), Color.FromArgb(248, 253, 147) };
        private static int _nextDefaultColorIndex = 0;
        private Panel? _dropIndicatorPanel;
        private TreeNode? _dropTargetNode;
        private bool _dropInsertBefore;
        private ProgressBar? _loadProgressBar;
        private Panel? _loadProgressPanel;
        private Label? _loadProgressLabel;
        private ToolStripMenuItem? _checkForUpdatesItem;
        private readonly HashSet<string> _dirtyFilePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _savedContentByFilePath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Stack<List<(int row, int col, string value)>>> _undoStackByFilePath = new Dictionary<string, Stack<List<(int row, int col, string value)>>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<EditHistoryItem>> _editHistoryByFilePath = new Dictionary<string, List<EditHistoryItem>>(StringComparer.OrdinalIgnoreCase);
        private (int row, int col, string value)? _pendingUndoCell;
        private bool _isApplyingUndo;
        private SplitContainer? _workspaceSplitContainer;
        private Panel? _editHistoryPanel;
        private ListBox? _editHistoryListBox;
        private bool _editHistoryPaneVisible = true;
        private ToolStripMenuItem? _editHistoryMenuItem;
        private SplitContainer? _bottomSplitContainer;
        private Panel? _myNotesPanel;
        private Label? _myNotesLabel;
        private TextBox? _myNotesTextBox;
        private bool _myNotesPaneVisible = true;
        private ToolStripMenuItem? _myNotesMenuItem;
        private bool _workspacePaneVisible = true;
        private ToolStripMenuItem? _workspacePaneMenuItem;
        private bool _darkMode = true;
        private ToolStripMenuItem? _darkModeMenuItem;
        private ToolStripMenuItem? _adaptiveColumnWidthMenuItem;
        private ToolStripMenuItem? _groupedColumnWidthMenuItem;
        private bool _inGroupedColumnWidthSync;
        private readonly Dictionary<TabPage, int[]> _savedColumnWidthsByTab = new Dictionary<TabPage, int[]>();
        private readonly Dictionary<TabPage, int[]> _savedColumnWidthsBeforeGroupedByTab = new Dictionary<TabPage, int[]>();
        private ToolStripMenuItem? _groupedRowHeightMenuItem;
        private bool _inGroupedRowHeightSync;
        private readonly Dictionary<TabPage, int[]> _savedRowHeightsBeforeGroupedByTab = new Dictionary<TabPage, int[]>();
        private Dictionary<string, int[]> _savedColumnWidthsByFile = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, int[]> _savedRowHeightsByFile = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
        private bool _autoLoadPreviousSession = true;
        private bool _preserveD2CompareWorkspaces = false;
        private ToolStripMenuItem? _autoLoadPreviousSessionMenuItem;
        private ToolStripMenuItem? _preserveD2CompareWorkspacesMenuItem;
        private ToolStripMenuItem? _headerTooltipsFromDataGuideMenuItem;
        private bool _headerTooltipsFromDataGuide;
        private bool _alwaysLockHeaderColumn;
        private bool _alwaysLockHeaderRow;
        private ToolStripMenuItem? _alwaysLockHeaderColumnMenuItem;
        private ToolStripMenuItem? _alwaysLockHeaderRowMenuItem;
        private Font? _gridFont;
        private bool _createBackupOnSave;
        private int _backupFileCount = 3;
        private int _contextMenuRestoreMode; // 0 none, 1 column, 2 row
        // Manual header tooltips: (fileName -> (columnName -> tooltip text)). Case-insensitive.
        private static readonly Dictionary<string, Dictionary<string, string>> ManualHeaderTooltips = BuildManualHeaderTooltips();
        private (int col, string? text)? _lastHeaderTooltipShown;
        private Control? _lastHeaderTooltipControl;
        private System.Windows.Forms.Timer? _headerTooltipPollTimer;
        private ToolTip? _headerTooltip;
        private TableLayoutPanel? _headerTooltipPanel;
        private Label? _headerTooltipLabelBold;
        private Label? _headerTooltipLabelNormal;
        private Panel? _workspaceSidePanel;
        private Panel? _workspaceToolbar;
        private Label? _editHistoryLabel;
        private Panel? _mainPanel;
        private Panel? _searchPanel;
        private TextBox? _searchTextBox;
        private Button? _searchPrevButton;
        private Button? _searchNextButton;
        private Label? _searchStatusLabel;
        private CheckBox? _searchMatchCaseCheck;
        private CheckBox? _searchMatchWholeCellCheck;
        private List<(int row, int col)> _searchMatches = new List<(int row, int col)>();
        private int _searchCurrentIndex = -1;
        private Panel? _mathPanel;
        private TextBox? _mathTextBox;
        private Button? _mathAddBtn;
        private Button? _mathSubBtn;
        private Button? _mathDivBtn;
        private Button? _mathMulBtn;
        private Button? _mathRoundUpBtn;
        private Button? _mathRoundDownBtn;
        private Button? _mathFormulaBtn;
        private Button? _mathIncrementFillBtn;

        private static readonly Color DarkBack = Color.FromArgb(45, 45, 48);
        private static readonly Color DarkPanel = Color.FromArgb(37, 37, 38);
        private static readonly Color DarkFore = Color.FromArgb(220, 220, 220);
        private static readonly Color SearchMatchActiveBack = Color.FromArgb(173, 214, 255);

        private Color _gridHeaderColor = Color.FromArgb(58, 65, 88);
        private Color _gridFrozenColor = Color.FromArgb(149, 215, 251);
        private Color _gridEvenRowColor = Color.FromArgb(44, 44, 44);
        private Color _gridOddRowColor = Color.FromArgb(62, 62, 62);
        private Color _gridSearchHighlightColor = Color.FromArgb(255, 255, 200);
        private Color _gridExpansionRowColor = Color.FromArgb(138, 0, 21);
        private readonly List<Color> _customPaletteColors = new List<Color>();

        private Color _textColorTab = SystemColors.ControlText;
        private Color _textColorHeader = Color.Black;
        private Color _textColorCell = SystemColors.ControlText;
        private Color _textColorCellFrozen = SystemColors.ControlText;
        private Color _textColorWorkspace = SystemColors.ControlText;
        private Color _textColorEditHistory = SystemColors.ControlText;
        private Color _textColorWorkspaceEntry = SystemColors.ControlText;
        private Color _textColorWorkspaceFiles = SystemColors.ControlText;
        private Color _textColorMenuStandby = SystemColors.ControlText;
        private Color _textColorMenuActive = SystemColors.HighlightText;
        private Color _textColorButtons = SystemColors.ControlText;

        private Color _textColorTabDark = Color.FromArgb(0, 0, 0);
        private Color _textColorHeaderDark = Color.FromArgb(220, 220, 220);
        private Color _textColorCellDark = Color.FromArgb(220, 220, 220);
        private Color _textColorCellFrozenDark = Color.FromArgb(0, 0, 0);
        private Color _textColorWorkspaceDark = Color.FromArgb(220, 220, 220);
        private Color _textColorEditHistoryDark = Color.FromArgb(220, 220, 220);
        private Color _textColorWorkspaceEntryDark = Color.FromArgb(0, 0, 0);
        private Color _textColorWorkspaceFilesDark = Color.FromArgb(200, 200, 200);
        private Color _textColorMenuStandbyDark = Color.FromArgb(220, 220, 220);
        private Color _textColorMenuActiveDark = Color.FromArgb(30, 30, 30);
        private Color _textColorButtonsDark = Color.FromArgb(220, 220, 220);

        private Color TextColorTabActive => _darkMode ? _textColorTabDark : _textColorTab;
        private Color TextColorHeaderActive => _darkMode ? _textColorHeaderDark : _textColorHeader;
        private Color TextColorCellActive => _darkMode ? _textColorCellDark : _textColorCell;
        private Color TextColorCellFrozenActive => _darkMode ? _textColorCellFrozenDark : _textColorCellFrozen;
        private Color TextColorWorkspaceActive => _darkMode ? _textColorWorkspaceDark : _textColorWorkspace;
        private Color TextColorEditHistoryActive => _darkMode ? _textColorEditHistoryDark : _textColorEditHistory;
        private Color TextColorWorkspaceEntryActive => _darkMode ? _textColorWorkspaceEntryDark : _textColorWorkspaceEntry;
        private Color TextColorWorkspaceFilesActive => _darkMode ? _textColorWorkspaceFilesDark : _textColorWorkspaceFiles;
        private Color TextColorMenuStandbyActive => _darkMode ? _textColorMenuStandbyDark : _textColorMenuStandby;
        private Color TextColorMenuActiveActive => _darkMode ? _textColorMenuActiveDark : _textColorMenuActive;
        private Color TextColorButtonsActive => _darkMode ? _textColorButtonsDark : _textColorButtons;

        public D2Horadrim()
        {
            InitializeComponent();

            treeView = new TreeView();
            treeView.HideSelection = false; // keep selection visible even when TreeView loses focus
            treeView.NodeMouseDoubleClick += TreeView_NodeMouseDoubleClick;
            treeView.MouseMove += TreeView_MouseMove;

            splitContainer1.Dock = DockStyle.Fill;
            splitContainer1.Orientation = Orientation.Vertical;
            splitContainer1.SplitterDistance = 220;
            splitContainer1.Panel1MinSize = 40;
            splitContainer1.SplitterMoving += (s, e) =>
            {
                if (e.SplitX > MaxWorkspacePanelWidth) e.Cancel = true;
            };

            _workspaceSidePanel = new Panel { Dock = DockStyle.Fill };
            _workspaceToolbar = new Panel { Dock = DockStyle.Top, Height = 28 };
            var workspaceSidePanel = _workspaceSidePanel;
            var workspaceToolbar = _workspaceToolbar;

            var lblWorkspaceManager = new Label
            {
                Text = "Workspace Manager",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = false
            };

            var btnPanel = new Panel { Dock = DockStyle.Right, Width = 58, Height = 28 };
            var btnAddWorkspace = new Button
            {
                Width = 26,
                Height = 26,
                Top = 1,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                Text = ""
            };
            btnAddWorkspace.Tag = "+";
            btnAddWorkspace.Paint += CenteredSymbolButton_Paint;

            var btnRemoveWorkspace = new Button
            {
                Width = 26,
                Height = 26,
                Top = 1,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                Text = ""
            };
            btnRemoveWorkspace.Tag = "-";
            btnRemoveWorkspace.Paint += CenteredSymbolButton_Paint;

            btnAddWorkspace.Left = btnPanel.Width - 54;
            btnRemoveWorkspace.Left = btnPanel.Width - 28;

            btnAddWorkspace.Click += BtnAddWorkspace_Click;
            btnRemoveWorkspace.Click += BtnRemoveWorkspace_Click;

            btnPanel.Controls.Add(btnAddWorkspace);
            btnPanel.Controls.Add(btnRemoveWorkspace);

            workspaceToolbar.Controls.Add(lblWorkspaceManager);
            workspaceToolbar.Controls.Add(btnPanel);

            treeView.LabelEdit = true;
            treeView.AllowDrop = true;
            treeView.DrawMode = TreeViewDrawMode.OwnerDrawAll;
            treeView.DrawNode += TreeView_DrawNode;
            treeView.BeforeLabelEdit += TreeView_BeforeLabelEdit;
            treeView.AfterLabelEdit += TreeView_AfterLabelEdit;
            treeView.ItemDrag += TreeView_ItemDrag;
            treeView.DragOver += TreeView_DragOver;
            treeView.DragDrop += TreeView_DragDrop;
            treeView.DragLeave += TreeView_DragLeave;

            _dropIndicatorPanel = new Panel
            {
                BackColor = SystemColors.HotTrack,
                Height = 3,
                Visible = false,
                Parent = workspaceSidePanel
            };

            workspaceSidePanel.Controls.Add(treeView);
            workspaceSidePanel.Controls.Add(_dropIndicatorPanel);
            _dropIndicatorPanel.BringToFront();
            workspaceSidePanel.Controls.Add(workspaceToolbar);

            // ---- Progress UI (timerless) ----
            _loadProgressPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                Visible = false,
                Padding = new Padding(4, 4, 4, 4)
            };

            // Loading Label
            _loadProgressLabel = new Label
            {
                Text = "Loading 0 / 0 KB",
                Dock = DockStyle.Top,
                Height = 22,
                Width = 240,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 2, 2, 2)
            };

            // ProgressBar fills the rest
            _loadProgressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            _loadProgressPanel.Controls.Add(_loadProgressBar);
            _loadProgressPanel.Controls.Add(_loadProgressLabel);
            workspaceSidePanel.Controls.Add(_loadProgressPanel);

            splitContainer1.Panel1.Controls.Add(workspaceSidePanel);

            // Workspace context menu (root nodes)
            var workspaceContextMenu = new ContextMenuStrip();
            var renameItem = new ToolStripMenuItem("Rename workspace");
            renameItem.Click += (s, _) =>
            {
                if (treeView.SelectedNode?.Parent == null)
                    treeView.SelectedNode.BeginEdit();
            };

            var setColorItem = new ToolStripMenuItem("Set color...");
            setColorItem.Click += (s, _) => SetWorkspaceColor();

            workspaceContextMenu.Items.Add(renameItem);
            workspaceContextMenu.Items.Add(setColorItem);

            // File context menu (workspace file nodes)
            var fileContextMenu = new ContextMenuStrip();
            var compareFilesItem = new ToolStripMenuItem("Compare files...");
            compareFilesItem.Click += (s, _) => CompareWorkspaceFile();
            fileContextMenu.Items.Add(compareFilesItem);

            // On right-click, set which context menu the tree should show and select the node.
            // Using the tree's ContextMenuStrip (set here) avoids another menu (e.g. form's) replacing ours on first click.
            treeView.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Right) return;
                var node = treeView.GetNodeAt(e.X, e.Y);
                if (node == null) return;
                treeView.SelectedNode = node;
                treeView.ContextMenuStrip = node.Parent == null ? workspaceContextMenu : fileContextMenu;
            };

            tabControl.Dock = DockStyle.Fill;
            tabControl.Appearance = TabAppearance.FlatButtons;

            // Clear any previous controls in Panel2
            splitContainer1.Panel2.Controls.Clear();

            // Create the horizontal split container
            _workspaceSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Panel1MinSize = 60,       // minimum height for top panel (tabControl)
                Panel2MinSize = 150,      // minimum height for edit history
                FixedPanel = FixedPanel.Panel2,  // bottom panel fixed for default height
                SplitterWidth = 4,
            };

            // Top panel: search bar + tabControl
            _mainPanel = new Panel { Dock = DockStyle.Fill };
            _mainPanel.Controls.Add(tabControl);

            _searchPanel = new Panel { Dock = DockStyle.Top, Height = 64, Padding = new Padding(4, 2, 4, 2) };
            var searchMathTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Height = 56
            };
            searchMathTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300));
            searchMathTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            searchMathTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var searchContentPanel = new Panel { Dock = DockStyle.Fill, Height = 56 };
            var searchLabel = new Label { Text = "Find:", AutoSize = true, Location = new Point(4, 6) };
            _searchTextBox = new TextBox { Width = 168, Location = new Point(37, 4), Anchor = AnchorStyles.Left };
            _searchTextBox.TextChanged += SearchTextBox_TextChanged;
            _searchTextBox.KeyDown += SearchTextBox_KeyDown;
            _searchStatusLabel = new Label { Text = "0 of 0", AutoSize = true, Location = new Point(212, 6) };
            int searchRow2Y = 30;
            int searchRow2X = 14;
            int searchBtnGap = 4;
            _searchMatchCaseCheck = new CheckBox
            {
                Text = "",
                Appearance = Appearance.Button,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(28, 24),
                Location = new Point(searchRow2X, searchRow2Y),
                Checked = false,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Bold)
            };
            _searchMatchCaseCheck.Paint += SearchMatchCaseCheck_Paint;
            _searchMatchCaseCheck.CheckedChanged += SearchMatchCheck_CheckedChanged;
            _searchMatchWholeCellCheck = new CheckBox
            {
                Text = "",
                Appearance = Appearance.Button,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(28, 24),
                Location = new Point(searchRow2X + 28 + searchBtnGap, searchRow2Y),
                Checked = false,
                Font = new Font("Segoe UI", 10f)
            };
            _searchMatchWholeCellCheck.Paint += SearchMatchWholeCellCheck_Paint;
            _searchMatchWholeCellCheck.CheckedChanged += SearchMatchCheck_CheckedChanged;
            toolTip1.SetToolTip(_searchMatchCaseCheck, "Match case");
            toolTip1.SetToolTip(_searchMatchWholeCellCheck, "Match entire cell contents");
            _searchPrevButton = new Button { Text = "", Size = new Size(70, 24), Location = new Point(searchRow2X + 28 + searchBtnGap + 28 + searchBtnGap, searchRow2Y), FlatStyle = FlatStyle.Flat };
            _searchPrevButton.Tag = "Previous";
            _searchPrevButton.Paint += SearchButton_Paint;
            _searchPrevButton.Click += SearchPrevButton_Click;
            _searchNextButton = new Button { Text = "", Size = new Size(70, 24), Location = new Point(searchRow2X + 28 + searchBtnGap + 28 + searchBtnGap + 70 + searchBtnGap, searchRow2Y), FlatStyle = FlatStyle.Flat };
            _searchNextButton.Tag = "Next";
            _searchNextButton.Paint += SearchButton_Paint;
            _searchNextButton.Click += SearchNextButton_Click;
            searchContentPanel.Controls.Add(searchLabel);
            searchContentPanel.Controls.Add(_searchTextBox);
            searchContentPanel.Controls.Add(_searchStatusLabel);
            searchContentPanel.Controls.Add(_searchMatchCaseCheck);
            searchContentPanel.Controls.Add(_searchMatchWholeCellCheck);
            searchContentPanel.Controls.Add(_searchPrevButton);
            searchContentPanel.Controls.Add(_searchNextButton);
            searchMathTable.Controls.Add(searchContentPanel, 0, 0);

            _mathPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 0, 0, 0) };
            var mathLabel = new Label { Text = "Math:", AutoSize = true, Location = new Point(0, 6) };
            _mathTextBox = new TextBox { Width = 220, Location = new Point(40, 24), Anchor = AnchorStyles.Left };
            int mathBtnY = 29;
            int mathBtnX = 4;
            int mathBtnW = 28;
            int mathBtnH = 24;
            int mathGap = 4;
            _mathAddBtn = CreateMathIconButton("+", "Addition", mathBtnX, mathBtnY, mathBtnW, mathBtnH);
            _mathSubBtn = CreateMathIconButton("−", "Subtraction", mathBtnX + (mathBtnW + mathGap), mathBtnY, mathBtnW, mathBtnH);
            _mathDivBtn = CreateMathIconButton("÷", "Division", mathBtnX + 2 * (mathBtnW + mathGap), mathBtnY, mathBtnW, mathBtnH);
            _mathMulBtn = CreateMathIconButton("×", "Multiplication", mathBtnX + 3 * (mathBtnW + mathGap), mathBtnY, mathBtnW, mathBtnH);
            _mathRoundUpBtn = CreateMathIconButton("↑", "Round Up", mathBtnX + 4 * (mathBtnW + mathGap), mathBtnY, mathBtnW, mathBtnH);
            _mathRoundDownBtn = CreateMathIconButton("↓", "Round Down", mathBtnX + 5 * (mathBtnW + mathGap), mathBtnY, mathBtnW, mathBtnH);
            _mathFormulaBtn = CreateMathIconButton("fx", "Custom Formula", mathBtnX + 6 * (mathBtnW + mathGap), mathBtnY, 32, mathBtnH);
            _mathIncrementFillBtn = CreateMathIconButton("1..", "Increment Fill", mathBtnX + 6 * (mathBtnW + mathGap) + 36, mathBtnY, mathBtnW, mathBtnH);
            _mathAddBtn.Click += (_, _) => ApplyMathOperation(MathOp.Add);
            _mathSubBtn.Click += (_, _) => ApplyMathOperation(MathOp.Subtract);
            _mathDivBtn.Click += (_, _) => ApplyMathOperation(MathOp.Divide);
            _mathMulBtn.Click += (_, _) => ApplyMathOperation(MathOp.Multiply);
            _mathRoundUpBtn.Click += (_, _) => ApplyMathOperation(MathOp.RoundUp);
            _mathRoundDownBtn.Click += (_, _) => ApplyMathOperation(MathOp.RoundDown);
            _mathFormulaBtn.Click += (_, _) => ApplyMathOperation(MathOp.CustomFormula);
            _mathIncrementFillBtn.Click += (_, _) => ApplyMathOperation(MathOp.IncrementFill);
            _mathPanel.Controls.Add(mathLabel);
            _mathPanel.Controls.Add(_mathTextBox);
            _mathPanel.Controls.Add(_mathAddBtn);
            _mathPanel.Controls.Add(_mathSubBtn);
            _mathPanel.Controls.Add(_mathDivBtn);
            _mathPanel.Controls.Add(_mathMulBtn);
            _mathPanel.Controls.Add(_mathRoundUpBtn);
            _mathPanel.Controls.Add(_mathRoundDownBtn);
            _mathPanel.Controls.Add(_mathFormulaBtn);
            _mathPanel.Controls.Add(_mathIncrementFillBtn);
            toolTip1.SetToolTip(_mathAddBtn, "Addition");
            toolTip1.SetToolTip(_mathSubBtn, "Subtraction");
            toolTip1.SetToolTip(_mathDivBtn, "Division");
            toolTip1.SetToolTip(_mathMulBtn, "Multiplication");
            toolTip1.SetToolTip(_mathRoundUpBtn, "Round Up");
            toolTip1.SetToolTip(_mathRoundDownBtn, "Round Down");
            toolTip1.SetToolTip(_mathFormulaBtn, "Custom Formula (x = cell value)");
            toolTip1.SetToolTip(_mathIncrementFillBtn, "Increment Fill (same column, top to bottom)");
            searchMathTable.Controls.Add(_mathPanel, 1, 0);
            _searchPanel.Controls.Add(searchMathTable);
            _mainPanel.Controls.Add(_searchPanel);

            _workspaceSplitContainer.Panel1.Controls.Add(_mainPanel);

            // Bottom panel: Edit History
            _editHistoryPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(4)
            };

            // Label at top of Edit History panel
            _editHistoryLabel = new Label
            {
                Text = "Edit history",
                Dock = DockStyle.Top,
                Height = 22,
                AutoSize = false,
                Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            var editHistoryLabel = _editHistoryLabel;

            // ListBox for edit history (OwnerDraw to show column/row/old/new in bold blue)
            _editHistoryListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                Font = new Font(this.Font.FontFamily, 9F),
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 18
            };
            _editHistoryListBox.DrawItem += EditHistoryListBox_DrawItem;

            // Add controls in correct order: ListBox first, label last
            _editHistoryPanel.Controls.Add(_editHistoryListBox);
            _editHistoryPanel.Controls.Add(editHistoryLabel);

            // Bottom strip: vertical split — left = Edit History, right = My Notes
            _bottomSplitContainer = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                Panel1MinSize = 80,
                Panel2MinSize = 80,
                SplitterWidth = 4
            };
            _bottomSplitContainer.Panel1.Controls.Add(_editHistoryPanel);

            _myNotesPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(4) };
            _myNotesLabel = new Label
            {
                Text = "My Notes",
                Dock = DockStyle.Top,
                Height = 22,
                AutoSize = false,
                Font = new Font(this.Font.FontFamily, 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            _myNotesTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font(this.Font.FontFamily, 9F),
                AcceptsReturn = true,
                WordWrap = true
            };
            _myNotesPanel.Controls.Add(_myNotesTextBox);
            _myNotesPanel.Controls.Add(_myNotesLabel);
            _bottomSplitContainer.Panel2.Controls.Add(_myNotesPanel);

            _workspaceSplitContainer.Panel2.Controls.Add(_bottomSplitContainer);

            // Initially collapse the bottom strip; My Notes right panel collapsed until toggled
            _workspaceSplitContainer.Panel2Collapsed = true;
            _bottomSplitContainer.Panel2Collapsed = true;

            // Add split container to Panel2
            splitContainer1.Panel2.Controls.Add(_workspaceSplitContainer);

            // Adjust SplitterDistance safely after form is loaded; load persisted notes
            this.Shown += (s, e) =>
            {
                int bottomHeight = splitContainer1.Panel2.Height - 150;
                if (bottomHeight > 0)
                    _workspaceSplitContainer.SplitterDistance = bottomHeight;
                if (_bottomSplitContainer != null && _bottomSplitContainer.Width > 0 && !_bottomSplitContainer.Panel1Collapsed && !_bottomSplitContainer.Panel2Collapsed)
                    _bottomSplitContainer.SplitterDistance = _bottomSplitContainer.Width / 2;
                LoadMyNotes();

                if (!string.IsNullOrEmpty(_startupFileArgument))
                {
                    string path = _startupFileArgument;
                    string? role = _startupWorkspaceRole;
                    string? folder = _startupWorkspaceFolder;
                    _startupFileArgument = null;
                    _startupWorkspaceRole = null;
                    _startupWorkspaceFolder = null;
                    if (File.Exists(path))
                    {
                        EnsureStartupWorkspaceForFile(path, role, folder);
                        var (wsColor, wsFolder) = GetWorkspaceColorForFilePath(path);
                        LoadFile(path, wsColor, wsFolder);
                        SelectWorkspaceNodeForFile(path);
                    }
                }
            };

            Text = "D2 Horadrim";

            treeView.Dock = DockStyle.Fill;
            SetDoubleBuffered(treeView);

            tabControl.Height = 30;
            tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabControl.SizeMode = TabSizeMode.Normal;
            tabControl.PaintTab += TabControl_PaintTab;
            tabControl.MouseDown += TabControl_MouseDown;
            tabControl.SelectedIndexChanged += tabControl_SelectedIndexChanged;

            contextMenuStrip1.Opening += ContextMenuStrip1_Opening;
            contextMenuStrip1.Closed += ContextMenuStrip1_Closed;

            tabControl.ContextMenuStrip = contextMenuStrip1;
            ContextMenuStrip = contextMenuStrip1;

            addColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Add();
            insertColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Insert();
            hideColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Hide();
            removeColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Remove();
            unhideColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Unhide();
            cloneColumnsToolStripMenuItem.Click += (s, ev) => ColumnOptions_Clone();

            addRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Add();
            insertRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Insert();
            hideRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Hide();
            removeRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Remove();
            unhideRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Unhide();
            cloneRowsToolStripMenuItem.Click += (s, ev) => RowOptions_Clone();

            _darkModeMenuItem = new ToolStripMenuItem("Dark mode");
            _darkModeMenuItem.CheckOnClick = true;
            _darkModeMenuItem.Checked = true;
            _darkModeMenuItem.Click += DarkModeMenuItem_Click;
            viewToolStripMenuItem.DropDownItems.Add(_darkModeMenuItem);

            _editHistoryMenuItem = new ToolStripMenuItem("Edit History Pane");
            _editHistoryMenuItem.CheckOnClick = true;
            _editHistoryMenuItem.Checked = true;
            _editHistoryMenuItem.Click += EditHistoryMenuItem_Click;
            viewToolStripMenuItem.DropDownItems.Add(_editHistoryMenuItem);

            _myNotesMenuItem = new ToolStripMenuItem("My Notes Pane");
            _myNotesMenuItem.CheckOnClick = true;
            _myNotesMenuItem.Checked = true;
            _myNotesMenuItem.Click += MyNotesMenuItem_Click;
            viewToolStripMenuItem.DropDownItems.Add(_myNotesMenuItem);

            _workspacePaneMenuItem = new ToolStripMenuItem("Workspace Pane");
            _workspacePaneMenuItem.CheckOnClick = true;
            _workspacePaneMenuItem.Checked = true;
            _workspacePaneMenuItem.Click += WorkspacePaneMenuItem_Click;
            viewToolStripMenuItem.DropDownItems.Add(_workspacePaneMenuItem);

            _headerTooltipsFromDataGuideMenuItem = new ToolStripMenuItem("Dataguide Tooltips (Header)");
            _headerTooltipsFromDataGuideMenuItem.CheckOnClick = true;
            _headerTooltipsFromDataGuideMenuItem.Checked = false;
            _headerTooltipsFromDataGuideMenuItem.Click += HeaderTooltipsFromDataGuideMenuItem_Click;
            viewToolStripMenuItem.DropDownItems.Add(_headerTooltipsFromDataGuideMenuItem);

            _autoLoadPreviousSessionMenuItem = new ToolStripMenuItem("Auto-load previous session");
            _autoLoadPreviousSessionMenuItem.CheckOnClick = true;
            _autoLoadPreviousSessionMenuItem.Checked = true;
            _autoLoadPreviousSessionMenuItem.Click += AutoLoadPreviousSessionMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_autoLoadPreviousSessionMenuItem);

            _preserveD2CompareWorkspacesMenuItem = new ToolStripMenuItem("Preserve D2Compare Workspace");
            _preserveD2CompareWorkspacesMenuItem.CheckOnClick = true;
            _preserveD2CompareWorkspacesMenuItem.Checked = false;
            _preserveD2CompareWorkspacesMenuItem.Click += PreserveD2CompareWorkspacesMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_preserveD2CompareWorkspacesMenuItem);

            _alwaysLockHeaderColumnMenuItem = new ToolStripMenuItem("Always lock/freeze header column");
            _alwaysLockHeaderColumnMenuItem.CheckOnClick = true;
            _alwaysLockHeaderColumnMenuItem.Checked = false;
            _alwaysLockHeaderColumnMenuItem.Click += AlwaysLockHeaderColumnMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_alwaysLockHeaderColumnMenuItem);

            _alwaysLockHeaderRowMenuItem = new ToolStripMenuItem("Always lock/freeze header row");
            _alwaysLockHeaderRowMenuItem.CheckOnClick = true;
            _alwaysLockHeaderRowMenuItem.Checked = false;
            _alwaysLockHeaderRowMenuItem.Click += AlwaysLockHeaderRowMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_alwaysLockHeaderRowMenuItem);

            _adaptiveColumnWidthMenuItem = new ToolStripMenuItem("Adaptive column width");
            _adaptiveColumnWidthMenuItem.CheckOnClick = true;
            _adaptiveColumnWidthMenuItem.Checked = false;
            _adaptiveColumnWidthMenuItem.Click += AdaptiveColumnWidthMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_adaptiveColumnWidthMenuItem);

            _groupedColumnWidthMenuItem = new ToolStripMenuItem("Grouped column width");
            _groupedColumnWidthMenuItem.CheckOnClick = true;
            _groupedColumnWidthMenuItem.Checked = false;
            _groupedColumnWidthMenuItem.Click += GroupedColumnWidthMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_groupedColumnWidthMenuItem);

            _groupedRowHeightMenuItem = new ToolStripMenuItem("Grouped row height");
            _groupedRowHeightMenuItem.CheckOnClick = true;
            _groupedRowHeightMenuItem.Checked = false;
            _groupedRowHeightMenuItem.Click += GroupedRowHeightMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(_groupedRowHeightMenuItem);

            var colorControlsItem = new ToolStripMenuItem("Background Color Controls");
            colorControlsItem.Click += ColorControlsMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(colorControlsItem);

            var textColorControlsItem = new ToolStripMenuItem("Text Color Controls");
            textColorControlsItem.Click += TextColorControlsMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(textColorControlsItem);

            var configurationItem = new ToolStripMenuItem("Configuration");
            configurationItem.Click += ConfigurationMenuItem_Click;
            settingsToolStripMenuItem.DropDownItems.Add(configurationItem);

            

            _headerTooltipPollTimer = new System.Windows.Forms.Timer(components) { Interval = 150 };
            _headerTooltipPollTimer.Tick += HeaderTooltipPollTimer_Tick;
            _headerTooltipPollTimer.Start();
            _headerTooltip = new ToolTip(components);
            _headerTooltip.InitialDelay = 0;
            _headerTooltip.AutoPopDelay = 10000;

            _headerTooltipLabelBold = new Label
            {
                AutoSize = false,
                BackColor = SystemColors.Info,
                ForeColor = SystemColors.InfoText,
                Font = new Font(SystemFonts.DefaultFont.FontFamily, SystemFonts.DefaultFont.SizeInPoints, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            _headerTooltipLabelNormal = new Label
            {
                AutoSize = false,
                Size = new Size(HeaderTooltipContentWidth, 1),
                BackColor = SystemColors.Info,
                ForeColor = SystemColors.InfoText,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            _headerTooltipPanel = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 2,
                BackColor = SystemColors.Info,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(HeaderTooltipPaddingH, HeaderTooltipPaddingV, HeaderTooltipPaddingH, HeaderTooltipPaddingV),
                Visible = false
            };
            _headerTooltipPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, HeaderTooltipContentWidth));
            _headerTooltipPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _headerTooltipPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _headerTooltipPanel.Controls.Add(_headerTooltipLabelBold, 0, 0);
            _headerTooltipPanel.Controls.Add(_headerTooltipLabelNormal, 0, 1);
            Controls.Add(_headerTooltipPanel);
            _headerTooltipPanel.BringToFront();

            var helpToolStripMenuItem = new ToolStripMenuItem("&Help");
            var hotkeysItem = new ToolStripMenuItem("&Hotkeys");
            hotkeysItem.Click += (_, _) => MessageBox.Show(HotkeysLegend, "Hotkeys", MessageBoxButtons.OK, MessageBoxIcon.Information);
            helpToolStripMenuItem.DropDownItems.Add(hotkeysItem);
            var openDataguideItem = new ToolStripMenuItem("Open D2R Dataguide (2.9)");
            openDataguideItem.Click += (_, _) => OpenD2RDataguide();
            helpToolStripMenuItem.DropDownItems.Add(openDataguideItem);
            _checkForUpdatesItem = new ToolStripMenuItem("Check for &updates");
            _checkForUpdatesItem.Click += (_, _) => CheckForUpdatesAsync();
            helpToolStripMenuItem.DropDownItems.Add(_checkForUpdatesItem);
            menuStrip1.Items.Add(helpToolStripMenuItem);
        }

        private const int EditHistoryPaneHeight = 140;

        private void HeaderTooltipGrid_MouseLeave(object? sender, EventArgs e)
        {
            if (_headerTooltipPanel != null)
            {
                _lastHeaderTooltipShown = null;
                _headerTooltipPanel.Visible = false;
            }
        }

        private DataGridView? GetGridUnderCursor()
        {
            Point screen = Cursor.Position;
            foreach (TabPage tab in tabControl.TabPages)
            {
                if (tab.Controls.Count == 0) continue;
                if (tab.Controls[0] is not DataGridView dgv) continue;
                if (!dgv.Visible) continue;
                Rectangle gridScreen = dgv.RectangleToScreen(dgv.ClientRectangle);
                if (gridScreen.Contains(screen)) return dgv;
            }
            return null;
        }

        private void HeaderTooltipPollTimer_Tick(object? sender, EventArgs e)
        {
            if (_headerTooltipPanel == null || _headerTooltipLabelBold == null || _headerTooltipLabelNormal == null) return;

            void HideTooltip()
            {
                _lastHeaderTooltipShown = null;
                _headerTooltipPanel.Visible = false;
            }

            if (!_headerTooltipsFromDataGuide)
            {
                if (_lastHeaderTooltipShown != null) HideTooltip();
                return;
            }
            DataGridView? dgv = GetGridUnderCursor();
            if (dgv == null || dgv.FindForm() == null)
            {
                if (_lastHeaderTooltipShown != null) HideTooltip();
                return;
            }
            Point client = dgv.PointToClient(Cursor.Position);
            var hit = dgv.HitTest(client.X, client.Y);
            bool isHeaderRow = hit.Type == DataGridViewHitTestType.ColumnHeader || hit.RowIndex == -1 || hit.RowIndex == 0;
            bool isHeader = hit.ColumnIndex >= 0 && isHeaderRow;
            if (!isHeader)
            {
                if (_lastHeaderTooltipShown != null) HideTooltip();
                return;
            }
            var col = dgv.Columns[hit.ColumnIndex];
            string columnName = (col.HeaderText ?? col.Name ?? "").Trim();
            if (string.IsNullOrEmpty(columnName) && hit.RowIndex == 0 && dgv.Rows.Count > 0 && !dgv.Rows[0].IsNewRow)
            {
                object? cellVal = dgv.Rows[0].Cells[hit.ColumnIndex].Value;
                columnName = (cellVal?.ToString() ?? "").Trim();
            }
            if (string.IsNullOrEmpty(columnName)) columnName = "(column " + hit.ColumnIndex + ")";
            if (columnName == OriginalIndexColumnName)
            {
                if (_lastHeaderTooltipShown != null) HideTooltip();
                return;
            }
            string? filePath = GetFilePathForGrid(dgv);
            string? fileName = filePath != null ? Path.GetFileName(filePath) : null;
            string? fromManual = fileName != null ? GetManualHeaderTooltip(fileName, columnName) : null;
            string tooltipText = !string.IsNullOrEmpty(fromManual) ? fromManual : "Column: " + columnName;
            if (_lastHeaderTooltipShown is (int c, string t) && c == hit.ColumnIndex && string.Equals(t, tooltipText, StringComparison.Ordinal))
                return;
            _lastHeaderTooltipShown = (hit.ColumnIndex, tooltipText);
            SetHeaderTooltipContent(_headerTooltipLabelBold, _headerTooltipLabelNormal, tooltipText);
            Point formPos = PointToClient(Cursor.Position);
            int x = Math.Max(4, Math.Min(formPos.X + 20, ClientSize.Width - 100));
            int y = Math.Max(4, Math.Min(formPos.Y + 25, ClientSize.Height - 50));
            _headerTooltipPanel.Location = new Point(x, y);
            _headerTooltipPanel.Visible = true;
            LayoutHeaderTooltipPanel();
            _headerTooltipPanel.BringToFront();
        }

        private void DarkModeMenuItem_Click(object? sender, EventArgs e)
        {
            _darkMode = _darkModeMenuItem?.Checked ?? false;
            ApplyTheme();
        }

        private void AdaptiveColumnWidthMenuItem_Click(object? sender, EventArgs e)
        {
            if (_adaptiveColumnWidthMenuItem?.Checked == true)
                ApplyAdaptiveColumnWidthToCurrentGrid();
            else
                RestoreSavedColumnWidths();
        }

        private void GroupedColumnWidthMenuItem_Click(object? sender, EventArgs e)
        {
            if (_groupedColumnWidthMenuItem?.Checked == true)
                SaveColumnWidthsBeforeGrouped();
            else
                RestoreSavedColumnWidthsBeforeGrouped();
            SaveWindowSettings();
        }

        private void GroupedRowHeightMenuItem_Click(object? sender, EventArgs e)
        {
            if (_groupedRowHeightMenuItem?.Checked == true)
                SaveRowHeightsBeforeGrouped();
            else
                RestoreSavedRowHeightsBeforeGrouped();
            SaveWindowSettings();
        }

        private void SaveColumnWidthsBeforeGrouped()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            var widths = new int[dgv.Columns.Count];
            for (int i = 0; i < dgv.Columns.Count; i++)
                widths[i] = dgv.Columns[i].Width;
            _savedColumnWidthsBeforeGroupedByTab[tab] = widths;
        }

        private void ApplyAdaptiveColumnWidthToCurrentGrid()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            // Save current column widths so we can restore when user unchecks
            var widths = new int[dgv.Columns.Count];
            for (int i = 0; i < dgv.Columns.Count; i++)
                widths[i] = dgv.Columns[i].Width;
            _savedColumnWidthsByTab[tab] = widths;
            _inGroupedColumnWidthSync = true;
            try
            {
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
            finally
            {
                _inGroupedColumnWidthSync = false;
            }
        }

        private void RestoreSavedColumnWidths()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            if (!_savedColumnWidthsByTab.TryGetValue(tab, out int[]? widths)) return;
            _inGroupedColumnWidthSync = true;
            try
            {
                for (int i = 0; i < dgv.Columns.Count && i < widths.Length; i++)
                    dgv.Columns[i].Width = widths[i];
            }
            finally
            {
                _inGroupedColumnWidthSync = false;
            }
        }

        private void DataGridView_ColumnWidthChanged(object? sender, DataGridViewColumnEventArgs e)
        {
            if (_groupedColumnWidthMenuItem?.Checked != true || _inGroupedColumnWidthSync) return;
            var dgv = sender as DataGridView;
            if (dgv == null || e.Column == null) return;
            int idx = e.Column.Index;
            int w = e.Column.Width;
            _inGroupedColumnWidthSync = true;
            try
            {
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (i != idx && dgv.Columns[i].Visible)
                        dgv.Columns[i].Width = w;
                }
            }
            finally
            {
                _inGroupedColumnWidthSync = false;
            }
        }

        private void RestoreSavedColumnWidthsBeforeGrouped()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            if (!_savedColumnWidthsBeforeGroupedByTab.TryGetValue(tab, out int[]? widths)) return;
            _inGroupedColumnWidthSync = true;
            try
            {
                for (int i = 0; i < dgv.Columns.Count && i < widths.Length; i++)
                    dgv.Columns[i].Width = widths[i];
            }
            finally
            {
                _inGroupedColumnWidthSync = false;
            }
        }

        private void SaveRowHeightsBeforeGrouped()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            var heights = new int[dgv.Rows.Count];
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].IsNewRow) continue;
                heights[i] = dgv.Rows[i].Height;
            }
            _savedRowHeightsBeforeGroupedByTab[tab] = heights;
        }

        private void RestoreSavedRowHeightsBeforeGrouped()
        {
            var tab = tabControl.SelectedTab;
            var dgv = GetCurrentGrid();
            if (dgv == null || tab == null) return;
            if (!_savedRowHeightsBeforeGroupedByTab.TryGetValue(tab, out int[]? heights)) return;
            _inGroupedRowHeightSync = true;
            try
            {
                for (int i = 0; i < dgv.Rows.Count && i < heights.Length; i++)
                {
                    if (dgv.Rows[i].IsNewRow) continue;
                    dgv.Rows[i].Height = heights[i];
                }
            }
            finally
            {
                _inGroupedRowHeightSync = false;
            }
        }

        private void DataGridView_RowHeightChanged(object? sender, DataGridViewRowEventArgs e)
        {
            if (_groupedRowHeightMenuItem?.Checked != true || _inGroupedRowHeightSync) return;
            var dgv = sender as DataGridView;
            if (dgv == null || e.Row == null || e.Row.IsNewRow) return;
            int h = e.Row.Height;
            _inGroupedRowHeightSync = true;
            try
            {
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    if (dgv.Rows[i].IsNewRow) continue;
                    dgv.Rows[i].Height = h;
                }
            }
            finally
            {
                _inGroupedRowHeightSync = false;
            }
        }

        private void AutoLoadPreviousSessionMenuItem_Click(object? sender, EventArgs e)
        {
            _autoLoadPreviousSession = _autoLoadPreviousSessionMenuItem?.Checked ?? true;
            SaveWindowSettings();
        }

        private void PreserveD2CompareWorkspacesMenuItem_Click(object? sender, EventArgs e)
        {
            _preserveD2CompareWorkspaces = _preserveD2CompareWorkspacesMenuItem?.Checked ?? false;
            SaveWindowSettings();
        }

        private void RunSearch()
        {
            _searchMatches.Clear();
            _searchCurrentIndex = -1;
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table)
            {
                UpdateSearchStatus();
                return;
            }
            string term = _searchTextBox?.Text?.Trim() ?? "";
            if (term.Length == 0)
            {
                UpdateSearchStatus();
                return;
            }
            bool matchCase = _searchMatchCaseCheck?.Checked ?? false;
            bool matchWholeCell = _searchMatchWholeCellCheck?.Checked ?? false;
            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            for (int r = 0; r < dgv.Rows.Count; r++)
            {
                if (dgv.Rows[r].IsNewRow) continue;
                for (int c = 0; c < dgv.Columns.Count; c++)
                {
                    if (dgv.Columns[c].Name == OriginalIndexColumnName) continue;
                    if (dgv.Columns[c].Visible == false) continue;
                    var val = dgv.Rows[r].Cells[c].Value?.ToString() ?? "";
                    bool isMatch = matchWholeCell ? string.Equals(val, term, comparison) : val.Contains(term, comparison);
                    if (isMatch)
                        _searchMatches.Add((r, c));
                }
            }
            _searchCurrentIndex = _searchMatches.Count > 0 ? 0 : -1;
            UpdateSearchStatus();
            if (_searchMatches.Count > 0)
                SelectSearchMatchAndScroll();
        }

        private void UpdateSearchStatus()
        {
            if (_searchStatusLabel == null) return;
            if (_searchMatches.Count == 0)
                _searchStatusLabel.Text = "0 of 0";
            else
                _searchStatusLabel.Text = $"{_searchCurrentIndex + 1} of {_searchMatches.Count}";
            if (_searchPrevButton != null) _searchPrevButton.Enabled = _searchMatches.Count > 0;
            if (_searchNextButton != null) _searchNextButton.Enabled = _searchMatches.Count > 0;
        }

        private void SelectSearchMatchAndScroll()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || _searchCurrentIndex < 0 || _searchCurrentIndex >= _searchMatches.Count) return;
            var (row, col) = _searchMatches[_searchCurrentIndex];
            dgv.ClearSelection();
            dgv.CurrentCell = dgv.Rows[row].Cells[col];
            dgv.Rows[row].Cells[col].Selected = true;
            dgv.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(row, dgv.RowCount - 1));
            if (dgv.Columns[col].Visible)
                dgv.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(col, dgv.ColumnCount - 1));
        }

        private void SearchTextBox_TextChanged(object? sender, EventArgs e)
        {
            RunSearch();
        }

        private void SearchTextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (e.Shift)
                    SearchPrevButton_Click(sender, e);
                else
                    SearchNextButton_Click(sender, e);
            }
        }

        private void SearchNextButton_Click(object? sender, EventArgs e)
        {
            if (_searchMatches.Count == 0) return;
            _searchCurrentIndex = (_searchCurrentIndex + 1) % _searchMatches.Count;
            UpdateSearchStatus();
            SelectSearchMatchAndScroll();
        }

        private void SearchPrevButton_Click(object? sender, EventArgs e)
        {
            if (_searchMatches.Count == 0) return;
            _searchCurrentIndex = _searchCurrentIndex <= 0 ? _searchMatches.Count - 1 : _searchCurrentIndex - 1;
            UpdateSearchStatus();
            SelectSearchMatchAndScroll();
        }

        private void ApplyTheme()
        {
            Color back = _darkMode ? DarkBack : SystemColors.Control;
            Color fore = _darkMode ? DarkFore : SystemColors.ControlText;
            Color panelBack = _darkMode ? DarkPanel : SystemColors.Control;

            BackColor = back;
            ForeColor = fore;
            splitContainer1.Panel1.BackColor = panelBack;
            splitContainer1.Panel2.BackColor = panelBack;
            tabPanel.BackColor = panelBack;
            menuStrip1.BackColor = panelBack;
            menuStrip1.ForeColor = fore;
            foreach (ToolStripItem item in menuStrip1.Items)
                item.ForeColor = fore;
            contextMenuStrip1.BackColor = panelBack;
            contextMenuStrip1.ForeColor = fore;
            foreach (ToolStripItem item in contextMenuStrip1.Items)
                item.ForeColor = fore;

            if (_workspaceSidePanel != null)
            {
                _workspaceSidePanel.BackColor = panelBack;
                foreach (Control c in _workspaceSidePanel.Controls)
                {
                    if (c == _dropIndicatorPanel)
                        c.BackColor = _darkMode ? Color.FromArgb(0, 122, 204) : SystemColors.HotTrack;
                    else
                        c.BackColor = panelBack;
                    if (c is TreeView tv) { tv.BackColor = panelBack; tv.ForeColor = fore; }
                    else if (c is Panel p)
                    {
                        p.BackColor = panelBack;
                        foreach (Control c2 in p.Controls)
                        {
                            c2.BackColor = panelBack;
                            c2.ForeColor = fore;
                        }
                    }
                }
            }
            if (_workspaceToolbar != null)
            {
                _workspaceToolbar.BackColor = panelBack;
                foreach (Control c in _workspaceToolbar.Controls)
                {
                    c.BackColor = panelBack;
                    c.ForeColor = fore;
                }
            }
            treeView.BackColor = panelBack;
            treeView.ForeColor = fore;

            if (_workspaceSplitContainer != null)
            {
                _workspaceSplitContainer.Panel1.BackColor = panelBack;
                _workspaceSplitContainer.Panel2.BackColor = panelBack;
            }
            if (_mainPanel != null)
                _mainPanel.BackColor = panelBack;
            tabControl.BackColor = panelBack;
            tabControl.TabStripBackColor = panelBack;
            if (_searchPanel != null)
            {
                _searchPanel.BackColor = panelBack;
                if (_searchMatchCaseCheck != null)
                    _searchMatchCaseCheck.BackColor = _searchMatchCaseCheck.Checked ? SearchMatchActiveBack : panelBack;
                if (_searchMatchWholeCellCheck != null)
                    _searchMatchWholeCellCheck.BackColor = _searchMatchWholeCellCheck.Checked ? SearchMatchActiveBack : panelBack;
            }
            if (_searchTextBox != null) { _searchTextBox.BackColor = back; _searchTextBox.ForeColor = fore; }
            if (_mathPanel != null) _mathPanel.BackColor = panelBack;
            if (_mathTextBox != null) { _mathTextBox.BackColor = back; _mathTextBox.ForeColor = fore; }
            if (_searchStatusLabel != null) _searchStatusLabel.ForeColor = fore;

            if (_editHistoryPanel != null)
            {
                _editHistoryPanel.BackColor = panelBack;
                foreach (Control c in _editHistoryPanel.Controls)
                {
                    c.BackColor = panelBack;
                    c.ForeColor = fore;
                }
            }
            if (_editHistoryListBox != null)
            {
                _editHistoryListBox.BackColor = panelBack;
                _editHistoryListBox.ForeColor = fore;
            }
            if (_editHistoryLabel != null)
            {
                _editHistoryLabel.BackColor = panelBack;
                _editHistoryLabel.ForeColor = fore;
            }
            if (_myNotesPanel != null)
            {
                _myNotesPanel.BackColor = panelBack;
                foreach (Control c in _myNotesPanel.Controls)
                {
                    c.BackColor = panelBack;
                    c.ForeColor = fore;
                }
            }
            if (_myNotesTextBox != null) { _myNotesTextBox.BackColor = back; _myNotesTextBox.ForeColor = fore; }
            if (_myNotesLabel != null) { _myNotesLabel.BackColor = panelBack; _myNotesLabel.ForeColor = fore; }

            if (_loadProgressPanel != null)
            {
                _loadProgressPanel.BackColor = panelBack;
                foreach (Control c in _loadProgressPanel.Controls)
                {
                    c.BackColor = panelBack;
                    if (c is Label || c is ProgressBar) c.ForeColor = fore;
                }
            }
            ApplyTextColors();
        }

        private void ApplyTextColors()
        {
            Color wsEntryColor = TextColorWorkspaceEntryActive;
            Color wsFilesColor = TextColorWorkspaceFilesActive;
            Color buttonColor = TextColorButtonsActive;
            Color ehColor = TextColorEditHistoryActive;
            Color menuStandby = TextColorMenuStandbyActive;
            Color menuActive = TextColorMenuActiveActive;

            treeView.ForeColor = wsFilesColor;
            foreach (TreeNode root in treeView.Nodes)
            {
                root.ForeColor = wsEntryColor;
                SetWorkspaceNodeColorsRecursive(root, wsEntryColor, wsFilesColor);
            }

            if (_workspaceToolbar != null)
            {
                foreach (Control c in _workspaceToolbar.Controls)
                {
                    c.ForeColor = buttonColor;
                    if (c is Panel p)
                    {
                        foreach (Control c2 in p.Controls)
                            c2.ForeColor = buttonColor;
                    }
                }
            }
            if (_workspaceSidePanel != null)
            {
                foreach (Control c in _workspaceSidePanel.Controls)
                {
                    if (c is TreeView) continue;
                    if (c is Panel p) { foreach (Control c2 in p.Controls) c2.ForeColor = buttonColor; }
                    else c.ForeColor = buttonColor;
                }
            }
            if (_searchPanel != null)
            {
                foreach (Control c in _searchPanel.Controls)
                {
                    c.ForeColor = buttonColor;
                    if (c is TableLayoutPanel tlp)
                    {
                        foreach (Control c2 in tlp.Controls)
                        {
                            c2.ForeColor = buttonColor;
                            if (c2 is Panel p)
                                foreach (Control c3 in p.Controls)
                                    c3.ForeColor = buttonColor;
                        }
                    }
                }
                _searchPrevButton?.Invalidate();
                _searchNextButton?.Invalidate();
            }
            if (_mathPanel != null)
            {
                foreach (Control c in _mathPanel.Controls)
                {
                    c.ForeColor = buttonColor;
                    if (c is Button) c.Invalidate();
                }
            }
            if (_editHistoryListBox != null)
                _editHistoryListBox.ForeColor = ehColor;
            if (_editHistoryLabel != null)
                _editHistoryLabel.ForeColor = ehColor;
            if (_myNotesTextBox != null)
                _myNotesTextBox.ForeColor = ehColor;
            if (_myNotesLabel != null)
                _myNotesLabel.ForeColor = ehColor;

            menuStrip1.ForeColor = menuStandby;
            foreach (ToolStripItem item in menuStrip1.Items)
                item.ForeColor = menuStandby;
            SetMenuStripRenderer(menuActive);

            contextMenuStrip1.ForeColor = menuStandby;
            foreach (ToolStripItem item in contextMenuStrip1.Items)
                item.ForeColor = menuStandby;
            contextMenuStrip1.Renderer = new MenuTextColorRenderer(menuActive);

            tabControl.Invalidate();
        }

        private void SetMenuStripRenderer(Color activeTextColor)
        {
            menuStrip1.Renderer = new MenuTextColorRenderer(activeTextColor);
        }

        private static void SetWorkspaceNodeColorsRecursive(TreeNode node, Color entryColor, Color filesColor)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.ForeColor = filesColor;
                SetWorkspaceNodeColorsRecursive(child, entryColor, filesColor);
            }
        }

        private void EditHistoryMenuItem_Click(object? sender, EventArgs e)
        {
            _editHistoryPaneVisible = _editHistoryMenuItem?.Checked ?? false;
            UpdateBottomStripVisibility();
            RefreshEditHistoryListBox();
        }

        private void MyNotesMenuItem_Click(object? sender, EventArgs e)
        {
            _myNotesPaneVisible = _myNotesMenuItem?.Checked ?? false;
            if (_bottomSplitContainer != null)
                _bottomSplitContainer.Panel2Collapsed = !_myNotesPaneVisible;
            UpdateBottomStripVisibility();
            if (_bottomSplitContainer != null && _myNotesPaneVisible && _editHistoryPaneVisible && _bottomSplitContainer.Width > 0)
                _bottomSplitContainer.SplitterDistance = _bottomSplitContainer.Width / 2;
        }

        private void WorkspacePaneMenuItem_Click(object? sender, EventArgs e)
        {
            _workspacePaneVisible = _workspacePaneMenuItem?.Checked ?? false;
            splitContainer1.Panel1Collapsed = !_workspacePaneVisible;
            SaveWindowSettings();
        }

        private void UpdateBottomStripVisibility()
        {
            if (_workspaceSplitContainer == null || _bottomSplitContainer == null) return;
            bool anyVisible = _editHistoryPaneVisible || _myNotesPaneVisible;
            _workspaceSplitContainer.Panel2Collapsed = !anyVisible;
            _bottomSplitContainer.Panel1Collapsed = !_editHistoryPaneVisible;
            _bottomSplitContainer.Panel2Collapsed = !_myNotesPaneVisible;
            if (anyVisible)
            {
                int total = _workspaceSplitContainer.Height - _workspaceSplitContainer.SplitterWidth;
                if (total > EditHistoryPaneHeight + 60)
                    _workspaceSplitContainer.SplitterDistance = total - EditHistoryPaneHeight;
            }
            if (_editHistoryPaneVisible && _myNotesPaneVisible && _bottomSplitContainer.Width > 0)
                _bottomSplitContainer.SplitterDistance = _bottomSplitContainer.Width / 2;
        }

        private List<EditHistoryItem> GetEditHistoryListForFilePath(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            if (string.IsNullOrEmpty(pathNorm)) pathNorm = "\0";
            if (!_editHistoryByFilePath.TryGetValue(pathNorm, out var list))
            {
                list = new List<EditHistoryItem>();
                _editHistoryByFilePath[pathNorm] = list;
            }
            return list;
        }

        private void RefreshEditHistoryListBox()
        {
            if (_editHistoryListBox == null) return;
            string? filePath = tabControl.SelectedTab?.Tag is TabTag tag ? tag.FilePath : null;
            _editHistoryListBox.Items.Clear();
            if (filePath != null)
            {
                var list = GetEditHistoryListForFilePath(filePath);
                foreach (var item in list)
                    _editHistoryListBox.Items.Add(item);
            }
        }

        private void ScrollEditHistoryToBottom()
        {
            if (_editHistoryListBox == null || _editHistoryListBox.Items.Count == 0) return;
            int visibleCount = Math.Max(1, _editHistoryListBox.Height / _editHistoryListBox.ItemHeight);
            _editHistoryListBox.TopIndex = Math.Max(0, _editHistoryListBox.Items.Count - visibleCount);
        }

        private void EditHistoryListBox_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || _editHistoryListBox == null) return;
            var item = _editHistoryListBox.Items[e.Index] as EditHistoryItem;
            e.DrawBackground();
            if (item == null) return;
            int x = e.Bounds.X + 2;
            var rect = new Rectangle(x, e.Bounds.Y, e.Bounds.Width - 2, e.Bounds.Height);
            using var normalFont = new Font(_editHistoryListBox.Font.FontFamily, _editHistoryListBox.Font.SizeInPoints);
            using var boldFont = new Font(normalFont, FontStyle.Bold);
            Color normalColor = (e.State & DrawItemState.Selected) != 0 ? SystemColors.HighlightText : (_editHistoryListBox.ForeColor);
            Color variableColor = (e.State & DrawItemState.Selected) != 0 ? SystemColors.HighlightText : Color.FromArgb(100, 180, 255);
            Color actionColor = (e.State & DrawItemState.Selected) != 0 ? SystemColors.HighlightText : Color.FromArgb(255, 120, 120);
            const TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix;

            int DrawSegment(string text, Font font, Color color)
            {
                if (string.IsNullOrEmpty(text)) return 0;
                var size = TextRenderer.MeasureText(e.Graphics, text, font, Size.Empty, flags);
                var segRect = new Rectangle(x, e.Bounds.Y, size.Width, e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, text, font, segRect, color, flags);
                x += size.Width;
                return size.Width;
            }

            string actionLabel = $"Action {e.Index + 1}: ";
            DrawSegment(actionLabel, boldFont, actionColor);

            if (item.ColumnName != null && item.RowName != null && item.OldValue != null && item.NewValue != null)
            {
                DrawSegment(item.ColumnName.Trim(), boldFont, variableColor);
                DrawSegment(" of ", normalFont, normalColor);
                DrawSegment(item.RowName.Trim(), boldFont, variableColor);
                DrawSegment(" edited from ", normalFont, normalColor);
                DrawSegment(item.OldValue.Trim(), boldFont, variableColor);
                DrawSegment(" to ", normalFont, normalColor);
                DrawSegment(item.NewValue.Trim(), boldFont, variableColor);
            }
            else if (item.PasteValue != null && item.PasteCount != null)
            {
                DrawSegment("Pasted the value ", normalFont, normalColor);
                DrawSegment(item.PasteValue.Trim(), boldFont, variableColor);
                DrawSegment(" into ", normalFont, normalColor);
                DrawSegment(item.PasteCount.Trim(), boldFont, variableColor);
                DrawSegment(" cell(s)", normalFont, normalColor);
                if (!string.IsNullOrEmpty(item.PasteCellRange))
                {
                    DrawSegment(" ", normalFont, normalColor);
                    DrawSegment(item.PasteCellRange.Trim(), boldFont, variableColor);
                }
            }
            else if (item.MathFormatKind != null && item.MathCellRange != null)
            {
                switch (item.MathFormatKind)
                {
                    case "Add":
                        DrawSegment(item.MathOperation!.Trim(), boldFont, variableColor);
                        DrawSegment(" ", normalFont, normalColor);
                        DrawSegment(item.MathValue!.Trim(), boldFont, variableColor);
                        DrawSegment(" to ", normalFont, normalColor);
                        DrawSegment(item.MathColumnName!.Trim(), boldFont, variableColor);
                        DrawSegment(" (", normalFont, normalColor);
                        DrawSegment(item.OldValue!.Trim(), boldFont, variableColor);
                        DrawSegment("->", normalFont, normalColor);
                        DrawSegment(item.NewValue!.Trim(), boldFont, variableColor);
                        DrawSegment(") for ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment("", normalFont, normalColor);
                        break;
                    case "Sub":
                        DrawSegment(item.MathOperation!.Trim(), boldFont, variableColor);
                        DrawSegment(" ", normalFont, normalColor);
                        DrawSegment(item.MathValue!.Trim(), boldFont, variableColor);
                        DrawSegment(" from ", normalFont, normalColor);
                        DrawSegment(item.MathColumnName!.Trim(), boldFont, variableColor);
                        DrawSegment(" (", normalFont, normalColor);
                        DrawSegment(item.OldValue!.Trim(), boldFont, variableColor);
                        DrawSegment("->", normalFont, normalColor);
                        DrawSegment(item.NewValue!.Trim(), boldFont, variableColor);
                        DrawSegment(") for ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment("", normalFont, normalColor);
                        break;
                    case "MulDiv":
                        DrawSegment(item.MathOperation!.Trim(), boldFont, variableColor);
                        DrawSegment(" ", normalFont, normalColor);
                        DrawSegment(item.MathColumnName!.Trim(), boldFont, variableColor);
                        DrawSegment(" by ", normalFont, normalColor);
                        DrawSegment(item.MathValue!.Trim(), boldFont, variableColor);
                        DrawSegment(" (", normalFont, normalColor);
                        DrawSegment(item.OldValue!.Trim(), boldFont, variableColor);
                        DrawSegment("->", normalFont, normalColor);
                        DrawSegment(item.NewValue!.Trim(), boldFont, variableColor);
                        DrawSegment(") for ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment("", normalFont, normalColor);
                        break;
                    case "Round":
                        DrawSegment(item.MathOperation!.Trim(), boldFont, variableColor);
                        DrawSegment(" ", normalFont, normalColor);
                        DrawSegment(item.MathColumnName!.Trim(), boldFont, variableColor);
                        DrawSegment(" ", normalFont, normalColor);
                        DrawSegment(item.MathRoundDirection!.Trim(), boldFont, variableColor);
                        DrawSegment(" (", normalFont, normalColor);
                        DrawSegment(item.OldValue!.Trim(), boldFont, variableColor);
                        DrawSegment("->", normalFont, normalColor);
                        DrawSegment(item.NewValue!.Trim(), boldFont, variableColor);
                        DrawSegment(") for ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment("", normalFont, normalColor);
                        break;
                    case "Custom":
                        DrawSegment("Applied ", normalFont, normalColor);
                        DrawSegment(item.MathValue!.Trim(), boldFont, variableColor);
                        DrawSegment(" to ", normalFont, normalColor);
                        DrawSegment(item.MathColumnName!.Trim(), boldFont, variableColor);
                        DrawSegment(" (", normalFont, normalColor);
                        DrawSegment(item.OldValue!.Trim(), boldFont, variableColor);
                        DrawSegment("->", normalFont, normalColor);
                        DrawSegment(item.NewValue!.Trim(), boldFont, variableColor);
                        DrawSegment(") for ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment("", normalFont, normalColor);
                        break;
                    case "IncrementFill":
                        DrawSegment("Incremented all values of ", normalFont, normalColor);
                        DrawSegment(item.MathCellRange.Trim(), boldFont, variableColor);
                        DrawSegment(" by 1", normalFont, normalColor);
                        break;
                    default:
                        DrawSegment(item.DisplayText, normalFont, normalColor);
                        break;
                }
            }
            else if (item.DisplayText.StartsWith("Math:", StringComparison.OrdinalIgnoreCase))
            {
                TextRenderer.DrawText(e.Graphics, item.DisplayText, boldFont, new Rectangle(x, e.Bounds.Y, e.Bounds.Width - (x - e.Bounds.X), e.Bounds.Height), variableColor, flags);
            }
            else
            {
                TextRenderer.DrawText(e.Graphics, item.DisplayText, normalFont, new Rectangle(x, e.Bounds.Y, e.Bounds.Width - (x - e.Bounds.X), e.Bounds.Height), normalColor, flags);
            }
            e.DrawFocusRectangle();
        }

        public bool PreFilterMessage(ref Message m)
        {
            // Block Alt key system messages so the menu bar is not activated; Alt+key combos still work (they use the other key's message)
            if ((m.Msg == WM_SYSKEYDOWN || m.Msg == WM_SYSKEYUP) && m.WParam.ToInt32() == VK_MENU)
                return true;
            return false;
        }

        private IntPtr LowLevelKeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
                return CallNextHookEx(_lowLevelHookHandle, nCode, wParam, lParam);

            IntPtr fg = GetForegroundWindow();
            if (fg == IntPtr.Zero || GetWindowThreadProcessId(fg, out uint pid) == 0 || pid != System.Diagnostics.Process.GetCurrentProcess().Id)
                return CallNextHookEx(_lowLevelHookHandle, nCode, wParam, lParam);

            var kbd = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
            int wp = wParam.ToInt32();

            if (wp == WM_SYSKEYDOWN)
            {
                if (kbd.vkCode == VK_MENU)
                    _altKeyLogicalDown = true;
                else
                    PostMessage(Handle, WM_HOTKEY_ALT, (IntPtr)kbd.vkCode, IntPtr.Zero);
                return (IntPtr)1;
            }
            if (wp == WM_SYSKEYUP)
            {
                if (kbd.vkCode == VK_MENU)
                    _altKeyLogicalDown = false;
                return (IntPtr)1;
            }
            if ((wp == WM_KEYDOWN || wp == WM_KEYUP) && kbd.vkCode == VK_MENU)
            {
                if (wp == WM_KEYDOWN)
                    _altKeyLogicalDown = true;
                else
                    _altKeyLogicalDown = false;
                return (IntPtr)1;
            }

            return CallNextHookEx(_lowLevelHookHandle, nCode, wParam, lParam);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == (int)WM_HOTKEY_ALT)
            {
                HandleAltHotkey(m.WParam.ToInt32());
                m.Result = IntPtr.Zero;
                return;
            }
            base.WndProc(ref m);
        }

        private static readonly int VK_A = 0x41, VK_C = 0x43, VK_D = 0x44, VK_E = 0x45, VK_F = 0x46, VK_H = 0x48, VK_I = 0x49, VK_N = 0x4E, VK_R = 0x52, VK_U = 0x55, VK_W = 0x57;

        private void HandleAltHotkey(int vkCode)
        {
            // Global toggles (work anywhere)
            if (vkCode == VK_E) { ToggleEditHistoryPane(); return; }
            if (vkCode == VK_N) { ToggleNotesPane(); return; }
            if (vkCode == VK_W) { ToggleWorkspacePane(); return; }
            if (vkCode == VK_D) { ToggleHeaderTooltips(); return; }

            // Row/column operations (require grid with selection)
            var dgv = GetCurrentGrid();
            if (dgv == null) return;
            bool columnOp = PreferColumnOperation(dgv);

            if (vkCode == VK_F)
            {
                ToggleFreezeSelectedColumnsOrRows(dgv, columnOp);
                return;
            }
            if (vkCode == VK_A) { if (columnOp) ColumnOptions_Add(); else RowOptions_Add(); return; }
            if (vkCode == VK_I) { if (columnOp) ColumnOptions_Insert(); else RowOptions_Insert(); return; }
            if (vkCode == VK_R) { if (columnOp) ColumnOptions_Remove(); else RowOptions_Remove(); return; }
            if (vkCode == VK_C) { if (columnOp) ColumnOptions_Clone(); else RowOptions_Clone(); return; }
            if (vkCode == VK_H) { if (columnOp) ColumnOptions_Hide(); else RowOptions_Hide(); return; }
            if (vkCode == VK_U) { if (columnOp) ColumnOptions_Unhide(); else RowOptions_Unhide(); return; }
        }

        private bool PreferColumnOperation(DataGridView dgv)
        {
            if (dgv.SelectedCells.Count == 0) return true;
            var cols = dgv.SelectedCells.Cast<DataGridViewCell>().Where(c => dgv.Columns[c.ColumnIndex].Name != OriginalIndexColumnName).Select(c => c.ColumnIndex).Distinct().Count();
            var rows = dgv.SelectedCells.Cast<DataGridViewCell>().Where(c => c.RowIndex >= 0 && !dgv.Rows[c.RowIndex].IsNewRow).Select(c => c.RowIndex).Distinct().Count();
            // Selection taller than wide (e.g. one column) -> column op; wider than tall (e.g. one row) -> row op
            return rows > cols;
        }

        /// <summary>
        /// Toggle freeze on the currently selected column(s) or row(s), using the same apply/refresh logic as the Settings "Always lock header column/row" and Alt+Click on header.
        /// </summary>
        private void ToggleFreezeSelectedColumnsOrRows(DataGridView dgv, bool columnOp)
        {
            if (dgv == null)
                return;

            dataGridView = dgv;

            if (columnOp)
            {
                // =========================
                // COLUMN LOCK / UNLOCK
                // =========================

                var colIndices = dgv.SelectedCells
                    .Cast<DataGridViewCell>()
                    .Select(c => c.ColumnIndex)
                    .Distinct()
                    .Where(c =>
                        c >= 0 &&
                        c < dgv.Columns.Count &&
                        dgv.Columns[c].Name != OriginalIndexColumnName)
                    .ToList();

                // Fallback to current cell if nothing selected
                if (colIndices.Count == 0 &&
                    dgv.CurrentCell != null &&
                    dgv.CurrentCell.ColumnIndex >= 0 &&
                    dgv.CurrentCell.ColumnIndex < dgv.Columns.Count &&
                    dgv.Columns[dgv.CurrentCell.ColumnIndex].Name != OriginalIndexColumnName)
                {
                    colIndices.Add(dgv.CurrentCell.ColumnIndex);
                }

                if (colIndices.Count == 0)
                    return;

                foreach (int c in colIndices)
                {
                    if (lockedColumns.Contains(c))
                        lockedColumns.Remove(c);
                    else
                        lockedColumns.Add(c);
                }

                // Row header is never lockable
                lockedColumns.Remove(-1);

                ApplyFrozenColumns();
                ApplyRowStyles(dgv);
                ApplyHeaderCellColors(dgv);
                RefreshColumnHeaders(dgv);
                dgv.Invalidate();
            }
            else
            {
                // =========================
                // ROW LOCK / UNLOCK
                // =========================

                lockedRows.Remove(-1); // never lock row header

                var rowIndices = dgv.SelectedCells
                    .Cast<DataGridViewCell>()
                    .Select(c => c.RowIndex)
                    .Distinct()
                    .Where(r =>
                        r >= 0 &&
                        r < dgv.Rows.Count &&
                        !dgv.Rows[r].IsNewRow)
                    .ToList();

                // Fallback to current row
                if (rowIndices.Count == 0 &&
                    dgv.CurrentCell != null &&
                    dgv.CurrentCell.RowIndex >= 0 &&
                    !dgv.Rows[dgv.CurrentCell.RowIndex].IsNewRow)
                {
                    rowIndices.Add(dgv.CurrentCell.RowIndex);
                }

                if (rowIndices.Count == 0)
                    return;

                foreach (int r in rowIndices)
                {
                    if (lockedRows.Contains(r))
                        lockedRows.Remove(r);
                    else
                        lockedRows.Add(r);
                }

                ApplyFrozenRows();
                ApplyHeaderCellColors(dgv);
                RefreshColumnHeaders(dgv);
                dgv.Invalidate();
            }
        }


        private void ToggleEditHistoryPane()
        {
            _editHistoryPaneVisible = !_editHistoryPaneVisible;
            if (_editHistoryMenuItem != null) _editHistoryMenuItem.Checked = _editHistoryPaneVisible;
            UpdateBottomStripVisibility();
            RefreshEditHistoryListBox();
        }

        private void ToggleNotesPane()
        {
            _myNotesPaneVisible = !_myNotesPaneVisible;
            if (_myNotesMenuItem != null) _myNotesMenuItem.Checked = _myNotesPaneVisible;
            if (_bottomSplitContainer != null) _bottomSplitContainer.Panel2Collapsed = !_myNotesPaneVisible;
            UpdateBottomStripVisibility();
            if (_bottomSplitContainer != null && _myNotesPaneVisible && _editHistoryPaneVisible && _bottomSplitContainer.Width > 0)
                _bottomSplitContainer.SplitterDistance = _bottomSplitContainer.Width / 2;
        }

        private void ToggleWorkspacePane()
        {
            _workspacePaneVisible = !_workspacePaneVisible;
            if (_workspacePaneMenuItem != null) _workspacePaneMenuItem.Checked = _workspacePaneVisible;
            splitContainer1.Panel1Collapsed = !_workspacePaneVisible;
            SaveWindowSettings();
        }

        private void ToggleHeaderTooltips()
        {
            _headerTooltipsFromDataGuide = !_headerTooltipsFromDataGuide;
            if (_headerTooltipsFromDataGuideMenuItem != null) _headerTooltipsFromDataGuideMenuItem.Checked = _headerTooltipsFromDataGuide;
            SaveWindowSettings();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (dataGridView?.Focused == true)
            {
                if (keyData == (Keys.Control | Keys.C))
                {
                    CopySelectedCellsToClipboard();
                    return true;
                }
                if (keyData == (Keys.Control | Keys.V))
                {
                    PasteFromClipboard();
                    return true;
                }
                if (keyData == (Keys.Control | Keys.Z))
                {
                    PerformUndo();
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private Stack<List<(int row, int col, string value)>> GetUndoStackForFilePath(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            if (string.IsNullOrEmpty(pathNorm)) pathNorm = "\0";
            if (!_undoStackByFilePath.TryGetValue(pathNorm, out var stack))
            {
                stack = new Stack<List<(int row, int col, string value)>>();
                _undoStackByFilePath[pathNorm] = stack;
            }
            return stack;
        }

        private void PerformUndo()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;
            if (tabControl.SelectedTab?.Tag is not TabTag tag || string.IsNullOrEmpty(tag.FilePath)) return;
            if (dataGridView.IsCurrentCellInEditMode)
                dataGridView.EndEdit();
            var stack = GetUndoStackForFilePath(tag.FilePath);
            if (stack.Count == 0) return;
            var entry = stack.Pop();
            var historyList = GetEditHistoryListForFilePath(tag.FilePath);
            if (historyList.Count > 0)
                historyList.RemoveAt(historyList.Count - 1);
            RefreshEditHistoryListBox();
            _isApplyingUndo = true;
            _pendingUndoCell = null;
            try
            {
                if (entry.Count == 1)
                {
                    var (r, c, value) = entry[0];
                    if (r == -1 && c == -1 && value != null && value.StartsWith("STRUCTURAL:", StringComparison.Ordinal))
                    {
                        PerformStructuralUndo(value, table);
                        SyncGridColumnDisplayOrderToTable(dataGridView, table);
                        dataGridView.Invalidate();
                        ApplyRowStyles(dataGridView);
                        ApplyHeaderCellColors(dataGridView);
                        UpdateColumnHeaders();
                        RefreshColumnHeaders();
                        return;
                    }
                }
                foreach (var (row, col, value) in entry)
                {
                    if (row < 0 || row >= dataGridView.Rows.Count || dataGridView.Rows[row].IsNewRow) continue;
                    if (col < 0 || col >= dataGridView.Columns.Count) continue;
                    if (dataGridView.Columns[col].Name == OriginalIndexColumnName) continue;
                    dataGridView.Rows[row].Cells[col].Value = value;
                }
                dataGridView.Invalidate();
            }
            finally
            {
                _isApplyingUndo = false;
            }
        }

        private const string StructuralPrefix = "STRUCTURAL:";

        private static void RenumberRowHeaderIndices(DataTable table)
        {
            if (!table.Columns.Contains(OriginalIndexColumnName)) return;
            for (int i = 0; i < table.Rows.Count; i++)
                table.Rows[i][OriginalIndexColumnName] = i;
        }

        private static void SyncGridColumnDisplayOrderToTable(DataGridView dgv, DataTable table)
        {
            if (dgv == null || table == null) return;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                string colName = table.Columns[i].ColumnName;
                if (!dgv.Columns.Contains(colName)) continue;
                dgv.Columns[colName].DisplayIndex = i;
            }
        }

        private void PerformStructuralUndo(string payload, DataTable table)
        {
            if (payload == null || !payload.StartsWith(StructuralPrefix, StringComparison.Ordinal)) return;
            string rest = payload.Substring(StructuralPrefix.Length);
            if (rest.StartsWith("REMOVE_COLUMNS:", StringComparison.Ordinal))
            {
                string names = rest.Substring("REMOVE_COLUMNS:".Length);
                char sep = '\x01';
                string[] colNames = names.Split(sep);
                foreach (string name in colNames)
                {
                    string n = name.Trim();
                    if (n.Length > 0 && table.Columns.Contains(n))
                        table.Columns.Remove(n);
                }
            }
            else if (rest.StartsWith("REMOVE_ROWS:", StringComparison.Ordinal))
            {
                string part = rest.Substring("REMOVE_ROWS:".Length);
                int colon = part.IndexOf(':');
                if (colon < 0) return;
                if (!int.TryParse(part.Substring(0, colon), out int startRow) || !int.TryParse(part.Substring(colon + 1), out int rowCount)) return;
                for (int i = 0; i < rowCount && startRow < table.Rows.Count; i++)
                    table.Rows.RemoveAt(startRow);
                RenumberRowHeaderIndices(table);
            }
            else if (rest.StartsWith("RESTORE_COLUMNS:", StringComparison.Ordinal))
            {
                string part = rest.Substring("RESTORE_COLUMNS:".Length);
                string[] segments = part.Split('\x02');
                if (segments.Length < 2) return;
                string[] ordNamePairs = segments[0].Split('\x01');
                var columnsToAdd = new List<(int ordinal, string name)>();
                foreach (string pair in ordNamePairs)
                {
                    int colon = pair.IndexOf(':');
                    if (colon < 0) continue;
                    if (!int.TryParse(pair.Substring(0, colon), out int ord)) continue;
                    string name = pair.Substring(colon + 1);
                    if (name.Length > 0 && !table.Columns.Contains(name))
                        columnsToAdd.Add((ord, name));
                }
                foreach (var (ordinal, name) in columnsToAdd.OrderBy(x => x.ordinal))
                {
                    var col = table.Columns.Add(name);
                    col.SetOrdinal(Math.Min(ordinal, table.Columns.Count - 1));
                }
                int numCols = columnsToAdd.Count;
                var colNames = columnsToAdd.OrderBy(x => x.ordinal).Select(x => x.name).ToList();
                for (int r = 0; r < table.Rows.Count && r + 1 < segments.Length; r++)
                {
                    string[] cells = segments[r + 1].Split('\x01');
                    for (int c = 0; c < numCols && c < cells.Length; c++)
                        table.Rows[r][colNames[c]] = UnescapeStructuralPayload(cells[c]);
                }
            }
            else if (rest.StartsWith("RESTORE_ROWS:", StringComparison.Ordinal))
            {
                string part = rest.Substring("RESTORE_ROWS:".Length);
                string[] segments = part.Split('\x02');
                if (segments.Length < 3) return;
                if (!int.TryParse(segments[0], out int rowCount)) return;
                string[] indicesStr = segments[1].Split('\x01');
                if (indicesStr.Length < rowCount) return;
                var indices = new List<int>();
                for (int i = 0; i < rowCount; i++)
                    if (int.TryParse(indicesStr[i], out int idx))
                        indices.Add(idx);
                int colCount = table.Columns.Count;
                for (int i = 0; i < indices.Count && i + 2 < segments.Length; i++)
                {
                    int insertAt = Math.Min(indices[i], table.Rows.Count);
                    var newRow = table.NewRow();
                    string[] cells = segments[i + 2].Split('\x01');
                    for (int c = 0; c < colCount && c < cells.Length; c++)
                        newRow[c] = UnescapeStructuralPayload(cells[c]);
                    table.Rows.InsertAt(newRow, insertAt);
                }
                RenumberRowHeaderIndices(table);
            }
        }

        private void PushStructuralUndo(string filePath, string payload, string description)
        {
            if (_isApplyingUndo) return;
            var entry = new List<(int row, int col, string value)> { (-1, -1, StructuralPrefix + payload) };
            GetUndoStackForFilePath(filePath).Push(entry);
            var historyItem = new EditHistoryItem(description, null, null, null, null, null, null, null, null, null, null, null, null, null);
            GetEditHistoryListForFilePath(filePath).Add(historyItem);
            RefreshEditHistoryListBox();
            ScrollEditHistoryToBottom();
        }

        private void PushUndoEntry(string filePath, List<(int row, int col, string value)> entry, string description, string? columnName = null, string? rowName = null, string? oldValue = null, string? newValue = null, string? pasteValue = null, string? pasteCount = null, string? pasteCellRange = null, string? mathOperation = null, string? mathValue = null, string? mathColumnName = null, string? mathCellRange = null, string? mathRoundDirection = null, string? mathFormatKind = null)
        {
            if (entry.Count == 0) return;
            if (_isApplyingUndo) return;
            GetUndoStackForFilePath(filePath).Push(entry);
            var historyItem = new EditHistoryItem(description, columnName, rowName, oldValue, newValue, pasteValue, pasteCount, pasteCellRange, mathOperation, mathValue, mathColumnName, mathCellRange, mathRoundDirection, mathFormatKind);
            GetEditHistoryListForFilePath(filePath).Add(historyItem);
            RefreshEditHistoryListBox();
            ScrollEditHistoryToBottom();
        }

        private string? GetFilePathForGrid(DataGridView? dgv)
        {
            if (dgv == null) return null;
            foreach (TabPage tab in tabControl.TabPages)
            {
                if (tab.Controls.Contains(dgv) && tab.Tag is TabTag tag)
                    return tag.FilePath;
            }
            return null;
        }

        private void DataGridView_KeyDown(object? sender, KeyEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null || dgv.DataSource is not DataTable) return;

            if (e.KeyCode == Keys.Delete)
            {
                if (dgv.SelectedCells.Count == 0) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                string? filePath = GetFilePathForGrid(dgv);
                var undoEntry = new List<(int row, int col, string value)>();
                foreach (DataGridViewCell cell in dgv.SelectedCells)
                {
                    if (cell.RowIndex < 0 || cell.RowIndex >= dgv.Rows.Count || dgv.Rows[cell.RowIndex].IsNewRow) continue;
                    if (cell.ColumnIndex < 0 || cell.ColumnIndex >= dgv.Columns.Count) continue;
                    if (dgv.Columns[cell.ColumnIndex].Name == OriginalIndexColumnName) continue;
                    undoEntry.Add((cell.RowIndex, cell.ColumnIndex, cell.Value?.ToString() ?? ""));
                }
                if (filePath != null && undoEntry.Count > 0)
                    PushUndoEntry(filePath, undoEntry, $"{undoEntry.Count} cell(s) cleared");
                foreach (DataGridViewCell cell in dgv.SelectedCells)
                {
                    if (cell.RowIndex < 0 || cell.RowIndex >= dgv.Rows.Count || dgv.Rows[cell.RowIndex].IsNewRow) continue;
                    if (cell.ColumnIndex < 0 || cell.ColumnIndex >= dgv.Columns.Count) continue;
                    if (dgv.Columns[cell.ColumnIndex].Name == OriginalIndexColumnName) continue;
                    cell.Value = "";
                }
                dgv.Invalidate();
            }
        }

        private void CopySelectedCellsToClipboard()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;
            if (dataGridView.SelectedCells.Count == 0) return;

            var selectedRows = dataGridView.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Where(r => r >= 0 && r < dataGridView.Rows.Count && !dataGridView.Rows[r].IsNewRow)
                .Distinct().OrderBy(r => r).ToList();
            var selectedCols = dataGridView.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.ColumnIndex)
                .Where(c => c >= 0 && c < dataGridView.Columns.Count && dataGridView.Columns[c].Name != OriginalIndexColumnName)
                .Distinct().OrderBy(c => c).ToList();

            if (selectedRows.Count == 0 || selectedCols.Count == 0) return;

            var sb = new StringBuilder();
            for (int ri = 0; ri < selectedRows.Count; ri++)
            {
                int r = selectedRows[ri];
                for (int ci = 0; ci < selectedCols.Count; ci++)
                {
                    if (ci > 0) sb.Append('\t');
                    int c = selectedCols[ci];
                    if (c < dataGridView.Rows[r].Cells.Count)
                        sb.Append(dataGridView.Rows[r].Cells[c].Value?.ToString() ?? "");
                }
                if (ri < selectedRows.Count - 1) sb.AppendLine();
            }

            try
            {
                Clipboard.SetText(sb.ToString(), TextDataFormat.UnicodeText);
            }
            catch { }
        }

        private void PasteFromClipboard()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table)
                return;
            string text;
            try
            {
                text = Clipboard.GetText(TextDataFormat.UnicodeText);
                if (string.IsNullOrWhiteSpace(text)) return;
            }
            catch { return; }

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            if (lines.Length == 0) return;

            bool singleCellPaste = lines.Length == 1 && lines[0].IndexOf('\t') < 0;
            if (singleCellPaste && dataGridView.SelectedCells.Count > 0)
            {
                string valueToPaste = lines[0];
                string? singleCellFilePath = tabControl.SelectedTab?.Tag is TabTag tab ? tab.FilePath : null;
                var singleCellUndoEntry = new List<(int row, int col, string value)>();
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.RowIndex < 0 || cell.RowIndex >= dataGridView.Rows.Count || dataGridView.Rows[cell.RowIndex].IsNewRow) continue;
                    if (cell.ColumnIndex < 0 || cell.ColumnIndex >= dataGridView.Columns.Count) continue;
                    if (dataGridView.Columns[cell.ColumnIndex].Name == OriginalIndexColumnName) continue;
                    singleCellUndoEntry.Add((cell.RowIndex, cell.ColumnIndex, cell.Value?.ToString() ?? ""));
                }
                if (singleCellFilePath != null && singleCellUndoEntry.Count > 0)
                {
                    string valuePreview = valueToPaste.Length > 50 ? valueToPaste.Substring(0, 47) + "..." : valueToPaste;
                    string? cellRange = FormatPasteCellRange(singleCellUndoEntry);
                    string desc = cellRange != null
                        ? $"Pasted the value {valuePreview} into {singleCellUndoEntry.Count} cell(s) {cellRange}"
                        : $"Pasted the value {valuePreview} into {singleCellUndoEntry.Count} cell(s)";
                    PushUndoEntry(singleCellFilePath, singleCellUndoEntry, desc,
                        pasteValue: valuePreview, pasteCount: singleCellUndoEntry.Count.ToString(), pasteCellRange: cellRange);
                }
                foreach (DataGridViewCell cell in dataGridView.SelectedCells)
                {
                    if (cell.RowIndex < 0 || cell.RowIndex >= dataGridView.Rows.Count || dataGridView.Rows[cell.RowIndex].IsNewRow) continue;
                    if (cell.ColumnIndex < 0 || cell.ColumnIndex >= dataGridView.Columns.Count) continue;
                    if (dataGridView.Columns[cell.ColumnIndex].Name == OriginalIndexColumnName) continue;
                    cell.Value = valueToPaste;
                }
                ApplyRowStyles();
                dataGridView.Invalidate();
                return;
            }

            int startRow;
            int startCol;
            if (dataGridView.SelectedCells.Count > 0)
            {
                var cells = dataGridView.SelectedCells.Cast<DataGridViewCell>().ToList();
                startRow = cells.Min(c => c.RowIndex);
                startCol = cells.Min(c => c.ColumnIndex);
            }
            else if (dataGridView.CurrentCell != null)
            {
                startRow = dataGridView.CurrentCell.RowIndex;
                startCol = dataGridView.CurrentCell.ColumnIndex;
            }
            else
            {
                startRow = 0;
                startCol = 0;
            }
            if (startRow < 0) startRow = 0;
            if (startCol < 0) startCol = 0;

            var dataColumnIndices = new List<int>();
            for (int c = 0; c < dataGridView.Columns.Count; c++)
            {
                if (dataGridView.Columns[c].Name != OriginalIndexColumnName)
                    dataColumnIndices.Add(c);
            }
            if (dataColumnIndices.Count == 0) return;

            int startColDataIndex = dataColumnIndices.IndexOf(startCol);
            if (startColDataIndex < 0) startColDataIndex = 0;

            int maxOrig = -1;
            int origColIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (origColIndex >= 0)
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    var v = table.Rows[r][OriginalIndexColumnName];
                    if (v != null && v != DBNull.Value && int.TryParse(v.ToString(), out int o) && o > maxOrig)
                        maxOrig = o;
                }
            }

            while (startRow + lines.Length > table.Rows.Count)
            {
                DataRow newRow = table.NewRow();
                if (origColIndex >= 0)
                    newRow[origColIndex] = ++maxOrig;
                table.Rows.Add(newRow);
            }

            string? filePath = tabControl.SelectedTab?.Tag is TabTag pt ? pt.FilePath : null;
            var pasteUndoEntry = new List<(int row, int col, string value)>();
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                string[] cells = lines[lineIndex].Split('\t');
                int rowIndex = startRow + lineIndex;

                if (rowIndex < 0 || rowIndex >= dataGridView.Rows.Count || dataGridView.Rows[rowIndex].IsNewRow)
                    continue;

                for (int cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                {
                    int dataColIndex = startColDataIndex + cellIndex;
                    if (dataColIndex >= dataColumnIndices.Count) break;
                    int gridCol = dataColumnIndices[dataColIndex];
                    string currentVal = dataGridView.Rows[rowIndex].Cells[gridCol].Value?.ToString() ?? "";
                    pasteUndoEntry.Add((rowIndex, gridCol, currentVal));
                }
            }
            if (filePath != null && pasteUndoEntry.Count > 0)
                PushUndoEntry(filePath, pasteUndoEntry, $"Paste ({pasteUndoEntry.Count} cell(s))");

            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
            {
                string[] cells = lines[lineIndex].Split('\t');
                int rowIndex = startRow + lineIndex;

                if (rowIndex < 0 || rowIndex >= dataGridView.Rows.Count || dataGridView.Rows[rowIndex].IsNewRow)
                    continue;

                for (int cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                {
                    int dataColIndex = startColDataIndex + cellIndex;
                    if (dataColIndex >= dataColumnIndices.Count) break;
                    int gridCol = dataColumnIndices[dataColIndex];
                    dataGridView.Rows[rowIndex].Cells[gridCol].Value = cellIndex < cells.Length ? cells[cellIndex] : "";
                }
            }

            ApplyRowStyles();
            dataGridView.Invalidate();
        }

        private void ApplyDefaultHeaderLocksIfEnabled()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;
            if (table.Rows.Count == 0) return;
            if (_alwaysLockHeaderColumn && dataGridView.Columns.Count > 0)
            {
                int firstDataCol = GetFirstDataColumnIndex(dataGridView);
                if (firstDataCol >= 0 && !lockedColumns.Contains(firstDataCol))
                    lockedColumns.Add(firstDataCol);
            }
            if (_alwaysLockHeaderRow && !lockedRows.Contains(0))
                lockedRows.Add(0);
        }

        private void UpdateColumnHeaders()
        {
            if (dataGridView == null) return;
            if (dataGridView.DataSource is not DataTable table || table.Rows.Count == 0)
                return;
            DataRow row0 = table.Rows[0];
            foreach (DataGridViewColumn col in dataGridView.Columns)
            {
                if (col.Name == OriginalIndexColumnName) continue;
                if (table.Columns.Contains(col.Name))
                    col.HeaderText = row0[col.Name]?.ToString() ?? "";
            }
        }

        private void SyncHeaderRowFromRow0()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table || table.Rows.Count == 0)
                return;
            DataRow row0 = table.Rows[0];
            int dataColCount = table.Columns.Count - (table.Columns.Contains(OriginalIndexColumnName) ? 1 : 0);
            var headerRow = new string[dataColCount];
            for (int c = 0; c < dataColCount; c++)
                headerRow[c] = row0[c]?.ToString() ?? "";
            table.ExtendedProperties["HeaderRow"] = headerRow;
        }

        private string GetColumnLetter(int index)
        {
            string columnLetter = "";
            while (index >= 0)
            {
                columnLetter = (char)('A' + (index % 26)) + columnLetter;
                index = (index / 26) - 1;
            }
            return columnLetter;
        }

        private string? FormatPasteCellRange(List<(int row, int col, string value)> entries)
        {
            if (dataGridView == null || entries.Count == 0) return null;
            int DataColIndex(int gridCol)
            {
                int d = 0;
                for (int c = 0; c < gridCol; c++)
                    if (dataGridView.Columns[c].Name != OriginalIndexColumnName) d++;
                return d;
            }
            var cells = entries.Select(e => (e.row, dataCol: DataColIndex(e.col))).Distinct().OrderBy(t => t.row).ThenBy(t => t.dataCol).ToList();
            if (cells.Count == 0) return null;
            // Row numbers are 0-based to match row header labels (grid row 0 -> "0", grid row 6 -> "6")
            int DisplayRow(int r) => r;
            if (cells.Count == 1)
                return $"({GetColumnLetter(cells[0].dataCol)}{DisplayRow(cells[0].row)})";
            int minR = cells.Min(t => t.row);
            int maxR = cells.Max(t => t.row);
            int minC = cells.Min(t => t.dataCol);
            int maxC = cells.Max(t => t.dataCol);
            int rectCount = (maxR - minR + 1) * (maxC - minC + 1);
            bool isRect = rectCount == cells.Count;
            if (isRect)
            {
                var set = new HashSet<(int row, int dataCol)>(cells);
                for (int r = minR; r <= maxR && isRect; r++)
                    for (int c = minC; c <= maxC; c++)
                        if (!set.Contains((r, c))) { isRect = false; break; }
                if (isRect)
                    return $"({GetColumnLetter(minC)}{DisplayRow(minR)}:{GetColumnLetter(maxC)}{DisplayRow(maxR)})";
            }
            const int maxList = 10;
            var refs = cells.Take(maxList).Select(t => $"{GetColumnLetter(t.dataCol)}{DisplayRow(t.row)}");
            string list = string.Join(", ", refs);
            if (cells.Count > maxList)
                list += $" and {cells.Count - maxList} more";
            return $"({list})";
        }

        private void DataGridView_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var dgv = (DataGridView)sender;
            if (dgv == null || e.RowIndex != -1 || e.ColumnIndex < 0) return;
            dataGridView = dgv;

            int columnIndex = e.ColumnIndex;
            dgv.Columns[columnIndex].SortMode = DataGridViewColumnSortMode.NotSortable;

            if (_altKeyLogicalDown || (Control.ModifierKeys & Keys.Alt) == Keys.Alt)
            {
                if (lockedColumns.Contains(columnIndex))
                    lockedColumns.Remove(columnIndex);
                else
                    lockedColumns.Add(columnIndex);
                ApplyFrozenColumns();
                ApplyRowStyles();
                ApplyHeaderCellColors(dgv);
                RefreshColumnHeaders(dgv);
                dgv.Invalidate();
            }
            else
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    for (int r = 0; r < dataGridView.Rows.Count; r++)
                    {
                        if (dataGridView.Rows[r].IsNewRow) continue;
                        if (columnIndex < dataGridView.Rows[r].Cells.Count)
                            dataGridView.Rows[r].Cells[columnIndex].Selected = true;
                    }
                    _anchorColumnIndex = columnIndex;
                }
                else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    int from = _anchorColumnIndex >= 0 ? Math.Min(_anchorColumnIndex, columnIndex) : columnIndex;
                    int to = _anchorColumnIndex >= 0 ? Math.Max(_anchorColumnIndex, columnIndex) : columnIndex;
                    dataGridView.ClearSelection();
                    for (int c = from; c <= to; c++)
                    {
                        for (int r = 0; r < dataGridView.Rows.Count; r++)
                        {
                            if (dataGridView.Rows[r].IsNewRow) continue;
                            if (c < dataGridView.Rows[r].Cells.Count)
                                dataGridView.Rows[r].Cells[c].Selected = true;
                        }
                    }
                    _anchorColumnIndex = columnIndex;
                }
                else
                {
                    dataGridView.ClearSelection();
                    for (int r = 0; r < dataGridView.Rows.Count; r++)
                    {
                        if (dataGridView.Rows[r].IsNewRow) continue;
                        if (columnIndex < dataGridView.Rows[r].Cells.Count)
                            dataGridView.Rows[r].Cells[columnIndex].Selected = true;
                    }
                    _anchorColumnIndex = columnIndex;
                }
            }
        }

        private void DataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var dgv = (DataGridView)sender;
            if (dgv == null || e.RowIndex < 0) return;
            dataGridView = dgv;

            int rowIndex = e.RowIndex;
            if (dgv.Rows[rowIndex].IsNewRow) return;

            if (_altKeyLogicalDown || (Control.ModifierKeys & Keys.Alt) == Keys.Alt)
            {
                if (lockedRows.Contains(rowIndex))
                    lockedRows.Remove(rowIndex);
                else
                    lockedRows.Add(rowIndex);
                ApplyFrozenRows();
                dgv.Invalidate();
            }
            else
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    for (int c = 0; c < dataGridView.Columns.Count; c++)
                    {
                        if (c < dataGridView.Rows[rowIndex].Cells.Count)
                            dataGridView.Rows[rowIndex].Cells[c].Selected = true;
                    }
                    _anchorRowIndex = rowIndex;
                }
                else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    int from = _anchorRowIndex >= 0 ? Math.Min(_anchorRowIndex, rowIndex) : rowIndex;
                    int to = _anchorRowIndex >= 0 ? Math.Max(_anchorRowIndex, rowIndex) : rowIndex;
                    dataGridView.ClearSelection();
                    for (int r = from; r <= to; r++)
                    {
                        if (dataGridView.Rows[r].IsNewRow) continue;
                        for (int c = 0; c < dataGridView.Columns.Count; c++)
                        {
                            if (c < dataGridView.Rows[r].Cells.Count)
                                dataGridView.Rows[r].Cells[c].Selected = true;
                        }
                    }
                    _anchorRowIndex = rowIndex;
                }
                else
                {
                    dataGridView.ClearSelection();
                    for (int c = 0; c < dataGridView.Columns.Count; c++)
                    {
                        if (c < dataGridView.Rows[rowIndex].Cells.Count)
                            dataGridView.Rows[rowIndex].Cells[c].Selected = true;
                    }
                    _anchorRowIndex = rowIndex;
                }
            }
        }

        private void ApplyFrozenColumns()
        {
            if (dataGridView == null)
                return;

            var dgv = dataGridView;

            // Clear frozen state first
            foreach (DataGridViewColumn col in dgv.Columns)
                col.Frozen = false;

            // Sort locked columns by original index
            var frozenCols = lockedColumns
                .Where(i => i >= 0 && i < dgv.Columns.Count)
                .OrderBy(i => i)
                .ToList();

            // Move frozen columns to the left, contiguously
            int displayIndex = 0;
            foreach (int colIndex in frozenCols)
            {
                dgv.Columns[colIndex].DisplayIndex = displayIndex++;
            }

            // Move remaining columns after frozen block
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (!frozenCols.Contains(col.Index))
                    col.DisplayIndex = displayIndex++;
            }

            // Now freeze the leftmost block
            foreach (int colIndex in frozenCols)
                dgv.Columns[colIndex].Frozen = true;
        }

        private void ApplyFrozenRows()
        {
            if (lockedRows.Count == 0)
            {
                RestoreRowOrderToOriginal();
                return;
            }
            ReorderDataTableRowsPutLockedFirst();
        }

        private void RestoreRowOrderToOriginal()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;
            if (!table.Columns.Contains(OriginalIndexColumnName)) return;

            var dv = new DataView(table) { Sort = $"[{OriginalIndexColumnName}] ASC" };
            DataTable sortedTable = table.Clone();
            if (table.ExtendedProperties["HeaderRow"] is string[] headerRow)
                sortedTable.ExtendedProperties["HeaderRow"] = headerRow;
            foreach (DataRowView drv in dv)
                sortedTable.Rows.Add(drv.Row.ItemArray);

            string? filePath = (tabControl.SelectedTab?.Tag as TabTag)?.FilePath ?? tabControl.SelectedTab?.Tag as string;
            if (!string.IsNullOrEmpty(filePath))
                fileCache[NormalizeFilePath(filePath)] = sortedTable;

            dataGridView.DataSource = sortedTable;
            HideOriginalIndexColumn();
            UpdateColumnHeaders();
            ApplyFrozenColumns();
            ApplyRowStyles();
            ApplyHeaderCellColors(dataGridView);
            RefreshColumnHeaders();
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                if (dataGridView.Rows[i].IsNewRow) continue;
                dataGridView.Rows[i].Frozen = false;
            }
            dataGridView.Invalidate();
        }

        private void ReorderDataTableRowsPutLockedFirst()
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;

            var sortedLocked = lockedRows.Where(i => i >= 0 && i < table.Rows.Count).OrderBy(i => i).ToList();
            if (sortedLocked.Count == 0)
            {
                lockedRows.Clear();
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    if (dataGridView.Rows[i].IsNewRow) continue;
                    dataGridView.Rows[i].Frozen = false;
                }
                dataGridView.Invalidate();
                return;
            }
            var unlockedIndices = Enumerable.Range(0, table.Rows.Count).Where(i => i != 0 && !lockedRows.Contains(i)).ToList();
            var lockedExcludingHeader = sortedLocked.Where(i => i != 0).ToList();
            bool headerRowLocked = sortedLocked.Contains(0);
            var newOrder = new List<int> { 0 }; // Row 0 = header always first
            newOrder.AddRange(lockedExcludingHeader);
            newOrder.AddRange(unlockedIndices);

            DataTable newTable = table.Clone();
            if (table.ExtendedProperties["HeaderRow"] is string[] headerRow)
                newTable.ExtendedProperties["HeaderRow"] = headerRow;
            foreach (int i in newOrder)
                newTable.Rows.Add(table.Rows[i].ItemArray);

            dataGridView.DataSource = newTable;
            HideOriginalIndexColumn();
            UpdateColumnHeaders();
            ApplyFrozenColumns();
            // Only count row 0 as locked when the user explicitly locked it; do not auto-lock the header row.
            int lockedCount = (headerRowLocked ? 1 : 0) + lockedExcludingHeader.Count;
            int lockedStart = headerRowLocked ? 0 : 1;
            lockedRows = Enumerable.Range(lockedStart, lockedCount).ToHashSet();
            // Freeze exactly the rows that are in lockedRows (their indices after reorder) so they stay visible when scrolling down.
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                if (dataGridView.Rows[i].IsNewRow) continue;
                dataGridView.Rows[i].Frozen = lockedRows.Contains(i);
            }
            ApplyRowStyles();
            ApplyHeaderCellColors(dataGridView);
            RefreshColumnHeaders();
            dataGridView.Invalidate();
        }

        private void RefreshColumnHeaders()
        {
            RefreshColumnHeaders(dataGridView);
        }

        private void RefreshColumnHeaders(DataGridView? dgv)
        {
            if (dgv == null) return;
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (lockedColumns.Contains(column.Index))
                {
                    column.HeaderCell.Style.Padding = new Padding(0, 0, 20, 0);
                    column.HeaderCell.Tag = lockImage;
                }
                else
                {
                    column.HeaderCell.Style.Padding = new Padding(0);
                    column.HeaderCell.Tag = null;
                }
            }

            dgv.Invalidate();
        }

        private void DataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            var dgv = (DataGridView)sender;

            // COLUMN HEADERS (A, B, C…) — always header color; lock icon still shown for locked columns
            if (e.RowIndex == -1)
            {
                if (e.ColumnIndex < 0 || e.ColumnIndex >= dgv.Columns.Count)
                {
                    e.PaintBackground(e.ClipBounds, true);
                    e.Handled = true;
                    return;
                }

                bool isLocked = e.ColumnIndex > 0 && lockedColumns.Contains(e.ColumnIndex);
                Color back = _gridHeaderColor; // header labels never get frozen color
                Color fore = TextColorHeaderActive;

                using (var backBrush = new SolidBrush(back))
                    e.Graphics.FillRectangle(backBrush, e.CellBounds);

                string letter = GetColumnLetter(e.ColumnIndex);
                var font = e.CellStyle?.Font ?? SystemFonts.DefaultFont;

                using (var letterFont = new Font(font.FontFamily, 9f, FontStyle.Bold))
                using (var textBrush = new SolidBrush(fore))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    e.Graphics.DrawString(letter, letterFont, textBrush, e.CellBounds, sf);
                }

                if (isLocked)
                {
                    int padding = -10;
                    int imageX = e.CellBounds.Right - lockImage.Width - padding;
                    int imageY = e.CellBounds.Top + (e.CellBounds.Height - lockImage.Height) / 2 + 7;
                    e.Graphics.DrawImage(lockImage, imageX, imageY);
                }

                e.Handled = true;
                return;
            }

            // ROW HEADERS (0, 1, 2, 3…) — always header/expansion color; lock icon still shown for locked rows
            if (e.ColumnIndex == -1 && e.RowIndex >= 0)
            {
                int firstDataCol = GetFirstDataColumnIndex(dgv);
                bool expansionRow = IsExpansionRow(dgv, e.RowIndex, firstDataCol);

                bool isLocked = e.RowIndex > 0 && lockedRows.Contains(e.RowIndex);
                Color back = expansionRow ? _gridExpansionRowColor : _gridHeaderColor; // header labels never get frozen color
                Color fore = TextColorHeaderActive;

                using (var backBrush = new SolidBrush(back))
                    e.Graphics.FillRectangle(backBrush, e.CellBounds);

                string rowNumber = (e.RowIndex + 1).ToString();
                if (dgv.Columns.Contains(OriginalIndexColumnName))
                {
                    var val = dgv.Rows[e.RowIndex].Cells[OriginalIndexColumnName].Value;
                    if (val != null && int.TryParse(val.ToString(), out int origIdx))
                        rowNumber = (origIdx + 1).ToString();
                }

                using (var textBrush = new SolidBrush(fore))
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    e.Graphics.DrawString(rowNumber, e.CellStyle?.Font ?? SystemFonts.DefaultFont, textBrush, e.CellBounds, sf);
                }

                if (isLocked)
                {
                    int imageX = e.CellBounds.Right - lockImage.Width + 10;
                    int imageY = e.CellBounds.Top + 2;
                    e.Graphics.DrawImage(lockImage, imageX, imageY);
                }

                e.Handled = true;
                return;
            }

            // SEARCH HIGHLIGHTED CELLS
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && dgv == GetCurrentGrid() && _searchMatches.Count > 0 && _searchCurrentIndex >= 0 && _searchCurrentIndex < _searchMatches.Count)
            {
                var match = _searchMatches[_searchCurrentIndex];
                if (match.row == e.RowIndex && match.col == e.ColumnIndex)
                {
                    using (var backBrush = new SolidBrush(_gridSearchHighlightColor))
                        e.Graphics.FillRectangle(backBrush, e.CellBounds);

                    string val = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
                    Color fore = e.CellStyle?.ForeColor ?? TextColorCellActive;

                    using (var foreBrush = new SolidBrush(fore))
                    using (var sf = new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center })
                    {
                        var rect = new Rectangle(e.CellBounds.X + 2, e.CellBounds.Y, e.CellBounds.Width - 4,e.CellBounds.Height);
                        e.Graphics.DrawString(val, e.CellStyle?.Font ?? dgv.Font, foreBrush, rect, sf);
                    }

                    e.Handled = true;
                }
            }
        }


        private void DisableColumnSorting()
        {
            if (dataGridView == null) return;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void OpenMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Text Files (*.txt)|*.txt";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    currentFilePath = openFileDialog.FileName;
                    lastOpenedDirectory = Path.GetDirectoryName(currentFilePath);
                    SaveLastDirectory();
                    LoadFile(currentFilePath, onComplete: LoadDirectoryFiles);
                }
            }
        }

        private void OpenFolderMenuItem_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Select folders to load";

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string[] selectedDirectories = folderBrowserDialog.SelectedPath.Split(';');

                    foreach (var folder in selectedDirectories)
                    {
                        LoadFilesFromFolder(folder);
                    }
                }
            }
        }

        private const string D2CompareSourceWorkspaceName = "D2Compare (Source)";
        private static string D2CompareSourceWorkspaceKey => D2CompareSourceWorkspaceName.Replace(" ", "_");

        private void EnsureStartupWorkspaceForFile(string filePath, string? role, string? originFolder)
        {
            string effectiveFolder = !string.IsNullOrWhiteSpace(originFolder)
                ? originFolder
                : (Path.GetDirectoryName(filePath) ?? string.Empty);

            string roleText = string.IsNullOrWhiteSpace(role) ? "Source" : role!;
            string workspaceName = $"D2Compare ({roleText})";
            string key = workspaceName.Replace(" ", "_");
            TreeNode? workspaceNode = treeView.Nodes.ContainsKey(key) ? treeView.Nodes[key] : null;
            if (workspaceNode == null)
            {
                // Use a color that is not already used by any existing workspace.
                Color color = GetNextUnusedWorkspaceColor();
                var tag = new WorkspaceTag { FolderPath = effectiveFolder, BackColor = color };
                workspaceNode = new TreeNode(workspaceName) { Name = key, Tag = tag, BackColor = color };
                treeView.Nodes.Add(workspaceNode);

                if (Directory.Exists(effectiveFolder))
                {
                    foreach (var file in Directory.GetFiles(effectiveFolder, "*.txt"))
                        workspaceNode.Nodes.Add(new TreeNode(Path.GetFileName(file)) { Tag = file });
                }

                // Ensure the startup file node exists even if it is outside the origin folder
                string pathNorm = NormalizeFilePath(filePath);
                bool hasFile = false;
                foreach (TreeNode child in workspaceNode.Nodes)
                {
                    if (child.Tag is string tagPath && string.Equals(NormalizeFilePath(tagPath), pathNorm, StringComparison.OrdinalIgnoreCase))
                    {
                        hasFile = true;
                        break;
                    }
                }
                if (!hasFile)
                    workspaceNode.Nodes.Add(new TreeNode(Path.GetFileName(filePath)) { Tag = filePath });
                SaveWorkspacesConfig();
                ApplyTextColors();
            }
            else
            {
                if (workspaceNode.Tag is WorkspaceTag existingTag &&
                    string.IsNullOrEmpty(existingTag.FolderPath) &&
                    !string.IsNullOrEmpty(effectiveFolder))
                {
                    existingTag.FolderPath = effectiveFolder;
                    workspaceNode.Tag = existingTag;
                }

                string pathNorm = NormalizeFilePath(filePath);
                foreach (TreeNode child in workspaceNode.Nodes)
                    if (child.Tag is string tagPath && string.Equals(NormalizeFilePath(tagPath), pathNorm, StringComparison.OrdinalIgnoreCase))
                        return;
                workspaceNode.Nodes.Add(new TreeNode(Path.GetFileName(filePath)) { Tag = filePath });
                SaveWorkspacesConfig();
            }
        }

        private string GetNextWorkspaceName()
        {
            int idx = 1;
            while (treeView.Nodes.ContainsKey($"Workspace_{idx}"))
                idx++;
            return $"Workspace {idx}";
        }

        private static Color GetNextDefaultColor()
        {
            Color c = DefaultWorkspaceColors[_nextDefaultColorIndex % DefaultWorkspaceColors.Length];
            _nextDefaultColorIndex++;
            return c;
        }

        // Choose a workspace color that is not already used by any existing root workspace node.
        // Falls back to the default cycling behavior if all palette colors are in use.
        private Color GetNextUnusedWorkspaceColor()
        {
            var used = new HashSet<int>();
            foreach (TreeNode root in treeView.Nodes)
            {
                if (root.Tag is WorkspaceTag tag)
                    used.Add(tag.BackColor.ToArgb());
            }

            foreach (var c in DefaultWorkspaceColors)
            {
                if (!used.Contains(c.ToArgb()))
                    return c;
            }

            // All default colors are already used; reuse via normal cycling.
            return GetNextDefaultColor();
        }

        private static void CenteredSymbolButton_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not string symbol) return;
            var rect = new Rectangle(0, 0, btn.Width, btn.Height);
            using (var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            })
            using (var font = new Font(btn.Font.FontFamily, 14f, FontStyle.Bold))
            using (var brush = new SolidBrush(btn.ForeColor))
            {
                e.Graphics.DrawString(symbol, font, brush, rect, sf);
            }
        }

        private static void SearchButton_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Button btn || btn.Tag is not string text) return;
            var rect = new Rectangle(0, 0, btn.Width, btn.Height);
            using (var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            })
            using (var brush = new SolidBrush(btn.ForeColor))
            {
                e.Graphics.DrawString(text, btn.Font, brush, rect, sf);
            }
        }

        private static Button CreateMathIconButton(string symbol, string tooltip, int x, int y, int w, int h)
        {
            var btn = new Button
            {
                Text = "",
                Tag = symbol,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(w, h),
                Location = new Point(x, y)
            };
            btn.Paint += SearchButton_Paint;
            return btn;
        }

        private void ApplyMathOperation(MathOp op)
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable) return;
            string input = _mathTextBox?.Text?.Trim() ?? "";
            double operand = 0;
            if (op != MathOp.RoundUp && op != MathOp.RoundDown && op != MathOp.CustomFormula && op != MathOp.IncrementFill)
            {
                if (!double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out operand))
                    return;
            }
            if (op == MathOp.CustomFormula && string.IsNullOrWhiteSpace(input))
                return;
            var cells = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Where(c => c.RowIndex >= 0 && c.RowIndex < dgv.Rows.Count && c.ColumnIndex >= 0 && c.ColumnIndex < dgv.Columns.Count)
                .Where(c => dgv.Columns[c.ColumnIndex].Name != OriginalIndexColumnName)
                .Distinct()
                .ToList();
            if (cells.Count == 0) return;
            string? filePath = GetFilePathForGrid(dgv);
            var undoEntry = new List<(int row, int col, string value)>();
            if (op == MathOp.IncrementFill)
            {
                var byCol = cells.GroupBy(c => c.ColumnIndex);
                foreach (var colGroup in byCol)
                {
                    var ordered = colGroup.OrderBy(c => c.RowIndex).ToList();
                    string topVal = ordered[0].Value?.ToString() ?? "";
                    if (!double.TryParse(topVal, NumberStyles.Float, CultureInfo.InvariantCulture, out double startVal))
                        startVal = 0;
                    double v = startVal;
                    foreach (var cell in ordered)
                    {
                        string oldVal = cell.Value?.ToString() ?? "";
                        undoEntry.Add((cell.RowIndex, cell.ColumnIndex, oldVal));
                        cell.Value = v.ToString(CultureInfo.InvariantCulture);
                        v += 1;
                    }
                }
            }
            else
            {
                foreach (var cell in cells)
                {
                    string oldVal = cell.Value?.ToString() ?? "";
                    undoEntry.Add((cell.RowIndex, cell.ColumnIndex, oldVal));
                    if (!double.TryParse(oldVal, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                        continue;
                    double result = op switch
                    {
                        MathOp.Add => val + operand,
                        MathOp.Subtract => val - operand,
                        MathOp.Divide => operand == 0 ? val : val / operand,
                        MathOp.Multiply => val * operand,
                        MathOp.RoundUp => Math.Ceiling(val),
                        MathOp.RoundDown => Math.Floor(val),
                        MathOp.CustomFormula => EvaluateFormula(input, val),
                        _ => val
                    };
                    if (op == MathOp.CustomFormula && double.IsNaN(result))
                        continue;
                    cell.Value = result.ToString(CultureInfo.InvariantCulture);
                }
            }
            if (filePath != null && undoEntry.Count > 0)
            {
                var firstEntry = undoEntry.OrderBy(t => t.row).ThenBy(t => t.col).First();
                string sampleOld = firstEntry.value;
                string sampleNew = dgv.Rows[firstEntry.row].Cells[firstEntry.col].Value?.ToString() ?? "";
                var cols = undoEntry.Select(t => t.col).Distinct().OrderBy(c => c).ToList();
                int minRow = undoEntry.Min(t => t.row);
                int maxRow = undoEntry.Max(t => t.row);
                string columnName = cols.Count == 1 ? dgv.Columns[cols[0]].HeaderText : string.Join(", ", cols.Select(c => dgv.Columns[c].HeaderText));
                string cellRange = $"rows {minRow}-{maxRow}";

                string pastTense = op switch
                {
                    MathOp.Add => "Added",
                    MathOp.Subtract => "Subtracted",
                    MathOp.Divide => "Divided",
                    MathOp.Multiply => "Multiplied",
                    MathOp.RoundUp => "Rounded Up",
                    MathOp.RoundDown => "Rounded Down",
                    MathOp.CustomFormula => "Applied custom formula",
                    MathOp.IncrementFill => "Increment filled",
                    _ => op.ToString()
                };
                string mathValueStr = op switch
                {
                    MathOp.Add or MathOp.Subtract or MathOp.Divide or MathOp.Multiply => operand.ToString(CultureInfo.InvariantCulture),
                    MathOp.CustomFormula => input.Length > 20 ? input.Substring(0, 17) + "…" : input,
                    _ => "—"
                };
                string formatKind = op switch
                {
                    MathOp.Add => "Add",
                    MathOp.Subtract => "Sub",
                    MathOp.Multiply or MathOp.Divide => "MulDiv",
                    MathOp.RoundUp or MathOp.RoundDown => "Round",
                    MathOp.CustomFormula => "Custom",
                    MathOp.IncrementFill => "IncrementFill",
                    _ => "AddSub"
                };
                string desc = formatKind == "IncrementFill"
                    ? $"Incremented all values of {cellRange} by 1"
                    : $"{pastTense} {mathValueStr} to {columnName} - {sampleOld} -> {sampleNew} ({cellRange})";
                PushUndoEntry(filePath, undoEntry, desc,
                    oldValue: formatKind == "IncrementFill" ? null : sampleOld,
                    newValue: formatKind == "IncrementFill" ? null : sampleNew,
                    mathOperation: pastTense, mathValue: mathValueStr, mathColumnName: columnName, mathCellRange: cellRange,
                    mathRoundDirection: "", mathFormatKind: formatKind);
                _dirtyFilePaths.Add(NormalizeFilePath(filePath));
                UpdateDirtyIndicators(filePath);
            }
            dgv.Invalidate();
        }

        private static double EvaluateFormula(string formula, double x)
        {
            string xStr = x.ToString(CultureInfo.InvariantCulture);
            string expr = formula.Replace("x", xStr).Replace("X", xStr);
            for (int i = 0; i < expr.Length; i++)
            {
                char c = expr[i];
                if (char.IsDigit(c) || c == '.' || c == '+' || c == '-' || c == '*' || c == '/' || c == '(' || c == ')' || c == ' ' || c == 'e' || c == 'E')
                    continue;
                return double.NaN;
            }
            try
            {
                var dt = new DataTable();
                object? obj = dt.Compute(expr, "");
                return Convert.ToDouble(obj, CultureInfo.InvariantCulture);
            }
            catch
            {
                return double.NaN;
            }
        }

        private void SearchMatchCheck_CheckedChanged(object? sender, EventArgs e)
        {
            if (sender is CheckBox cb)
                cb.BackColor = cb.Checked ? SearchMatchActiveBack : (cb.Parent?.BackColor ?? SystemColors.Control);
            RunSearch();
        }

        private static void SearchMatchCaseCheck_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not CheckBox cb) return;
            var rect = new Rectangle(0, 0, cb.Width, cb.Height);
            using (var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            })
            using (var brush = new SolidBrush(cb.ForeColor))
            {
                e.Graphics.DrawString("Aa", cb.Font, brush, rect, sf);
            }
        }

        private static void SearchMatchWholeCellCheck_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not CheckBox cb) return;
            var rect = new Rectangle(0, 0, cb.Width, cb.Height);
            using (var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            })
            using (var brush = new SolidBrush(cb.ForeColor))
            {
                e.Graphics.DrawString("▣", cb.Font, brush, rect, sf);
            }
        }

        private void BtnAddWorkspace_Click(object? sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog { Description = "Select folder for workspace" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                AddWorkspaceFromFolder(dlg.SelectedPath);
            }
        }

        private void BtnRemoveWorkspace_Click(object? sender, EventArgs e)
        {
            RemoveSelectedWorkspace();
        }

        private void RemoveSelectedWorkspace()
        {
            if (treeView.SelectedNode?.Parent != null) return;
            treeView.SelectedNode?.Remove();
            SaveWorkspacesConfig();
        }

        private static string NormalizeFolderPath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            path = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            try { return Path.GetFullPath(path); } catch { return path; }
        }

        private (Color? workspaceColor, string? workspaceFolderPath) GetWorkspaceColorForFilePath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return (null, null);
            string? fileDir = null;
            try { fileDir = Path.GetFullPath(Path.GetDirectoryName(filePath) ?? ""); } catch { return (null, null); }
            if (string.IsNullOrEmpty(fileDir)) return (null, null);
            string fileDirNorm = NormalizeFolderPath(fileDir);
            foreach (TreeNode node in treeView.Nodes)
            {
                if (node.Tag is not WorkspaceTag wsTag || string.IsNullOrEmpty(wsTag.FolderPath)) continue;
                string wsNorm = NormalizeFolderPath(wsTag.FolderPath);
                if (string.Equals(fileDirNorm, wsNorm, StringComparison.OrdinalIgnoreCase)
                    || (fileDirNorm.StartsWith(wsNorm + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)))
                    return (wsTag.BackColor, wsTag.FolderPath);
            }
            return (null, null);
        }

        private void SetWorkspaceColor()
        {
            if (treeView.SelectedNode?.Parent != null) return;
            if (treeView.SelectedNode?.Tag is not WorkspaceTag tag) return;
            using (var paletteForm = new ColorPaletteForm(tag.BackColor, _customPaletteColors) { StartPosition = FormStartPosition.CenterParent })
            {
                if (paletteForm.ShowDialog(this) != DialogResult.OK) return;
                SaveCustomPaletteColors();
                Color chosen = paletteForm.ChosenColor;
                tag.BackColor = chosen;
                treeView.SelectedNode.BackColor = chosen;
                SaveWorkspacesConfig();
                string folderPathNorm = NormalizeFolderPath(tag.FolderPath);
                foreach (TabPage tab in tabControl.TabPages)
                {
                    if (tab.Tag is TabTag t && string.Equals(NormalizeFolderPath(t.WorkspaceFolderPath), folderPathNorm, StringComparison.OrdinalIgnoreCase))
                    {
                        t.WorkspaceBackColor = chosen;
                        tab.BackColor = chosen;
                    }
                }
                tabControl.Invalidate();
            }
        }

        private static string GetWorkspacesConfigPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "D2_Horadrim");
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            return Path.Combine(dir, "workspaces.json");
        }

        private static string GetMyNotesFilePath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "D2_Horadrim");
            return Path.Combine(dir, "mynotes.txt");
        }

        private void LoadMyNotes()
        {
            if (_myNotesTextBox == null) return;
            try
            {
                string path = GetMyNotesFilePath();
                if (File.Exists(path))
                    _myNotesTextBox.Text = File.ReadAllText(path);
            }
            catch { }
        }

        private void SaveMyNotes()
        {
            if (_myNotesTextBox == null) return;
            try
            {
                string path = GetMyNotesFilePath();
                string? dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);
                File.WriteAllText(path, _myNotesTextBox.Text);
            }
            catch { }
        }

        private void SaveWorkspacesConfig()
        {
            try
            {
                var list = new List<WorkspaceEntry>();
                foreach (TreeNode node in treeView.Nodes)
                {
                    if (node.Tag is WorkspaceTag tag && !string.IsNullOrEmpty(tag.FolderPath))
                        list.Add(new WorkspaceEntry { Name = node.Text, FolderPath = tag.FolderPath, ColorArgb = tag.BackColor.ToArgb() });
                }
                var config = new WorkspacesConfig { Workspaces = list };
                string path = GetWorkspacesConfigPath();
                string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
            }
            catch { }
        }

        private void LoadWorkspacesConfig()
        {
            try
            {
                string path = GetWorkspacesConfigPath();
                if (!File.Exists(path)) return;

                string json = File.ReadAllText(path);
                var config = JsonSerializer.Deserialize<WorkspacesConfig>(json);
                if (config?.Workspaces == null || config.Workspaces.Count == 0) return;

                treeView.Nodes.Clear();
                for (int i = 0; i < config.Workspaces.Count; i++)
                {
                    var entry = config.Workspaces[i];
                    if (string.IsNullOrEmpty(entry.FolderPath)) continue;

                    Color color = entry.ColorArgb.HasValue ? Color.FromArgb(entry.ColorArgb.Value) : GetNextDefaultColor();
                    var tag = new WorkspaceTag { FolderPath = entry.FolderPath, BackColor = color };
                    string key = $"Workspace_{i + 1}";
                    TreeNode folderNode = new TreeNode(entry.Name) { Name = key, Tag = tag, BackColor = color };
                    treeView.Nodes.Add(folderNode);

                    if (Directory.Exists(entry.FolderPath))
                    {
                        foreach (var file in Directory.GetFiles(entry.FolderPath, "*.txt"))
                            folderNode.Nodes.Add(new TreeNode(Path.GetFileName(file)) { Tag = file });
                    }
                }
                ApplyTextColors();
            }
            catch { }
        }

        private void AddWorkspaceFromFolder(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath)) return;

            string name = GetNextWorkspaceName();
            Color color = GetNextDefaultColor();
            var tag = new WorkspaceTag { FolderPath = folderPath, BackColor = color };
            TreeNode folderNode = new TreeNode(name) { Name = name.Replace(" ", "_"), Tag = tag, BackColor = color };
            treeView.Nodes.Add(folderNode);

            foreach (var file in Directory.GetFiles(folderPath, "*.txt"))
                folderNode.Nodes.Add(new TreeNode(Path.GetFileName(file)) { Tag = file });
            SaveWorkspacesConfig();
            ApplyTextColors();
        }

        private void LoadFilesFromFolder(string folderPath)
        {
            AddWorkspaceFromFolder(folderPath);
        }

        private static string NormalizeFilePath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            try { return Path.GetFullPath(path); } catch { return path; }
        }

        private DataGridView? GetGridForTab(TabPage? tab)
        {
            if (tab == null) return null;
            foreach (Control c in tab.Controls)
                if (c is DataGridView dgv) return dgv;
            return null;
        }

        private DataGridView? GetCurrentGrid()
        {
            return GetGridForTab(tabControl.SelectedTab);
        }

        private void ColumnOptions_Add()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            string input = Microsoft.VisualBasic.Interaction.InputBox("How many columns to add?", "Add Columns", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;
            int insertBeforeIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (insertBeforeIndex < 0) insertBeforeIndex = table.Columns.Count;
            int dataColCount = table.Columns.Count - (table.Columns.Contains(OriginalIndexColumnName) ? 1 : 0);
            var newColNames = new List<string>();
            for (int i = 0; i < count; i++)
            {
                string colName = "Col" + (dataColCount + i);
                while (table.Columns.Contains(colName)) colName += "_";
                var col = table.Columns.Add(colName);
                col.SetOrdinal(insertBeforeIndex + i);
                newColNames.Add(colName);
            }
            for (int r = 0; r < table.Rows.Count; r++)
            {
                foreach (string colName in newColNames)
                    table.Rows[r][colName] = "";
            }
            MarkCurrentTabFileDirty();
            string? fp = GetFilePathForGrid(dgv);
            if (fp != null) PushStructuralUndo(fp, "REMOVE_COLUMNS:" + string.Join("\x01", newColNames), "Add columns");
            int firstNewColIndex = table.Columns.IndexOf(newColNames[0]);
            dgv.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(firstNewColIndex, dgv.ColumnCount - 1));
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
            UpdateColumnHeaders();
            RefreshColumnHeaders();
            ReapplyAllGridStylesDeferred(dgv);
        }

        private void ColumnOptions_Insert()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            int selectedGridCol = -1;
            if (dgv.SelectedCells.Count > 0)
            {
                selectedGridCol = dgv.SelectedCells.Cast<DataGridViewCell>()
                    .Where(c => dgv.Columns[c.ColumnIndex].Name != OriginalIndexColumnName)
                    .Select(c => c.ColumnIndex)
                    .DefaultIfEmpty(-1)
                    .Max();
            }
            if (selectedGridCol < 0 && dgv.CurrentCell != null && dgv.Columns[dgv.CurrentCell.ColumnIndex].Name != OriginalIndexColumnName)
                selectedGridCol = dgv.CurrentCell.ColumnIndex;
            if (selectedGridCol < 0)
            {
                MessageBox.Show("Select a column (or a cell in the column) after which to insert.", "Insert Columns", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string? selectedTableColName = dgv.Columns[selectedGridCol].DataPropertyName ?? dgv.Columns[selectedGridCol].Name;
            int selectedTableColIndex = table.Columns.IndexOf(selectedTableColName ?? "");
            if (selectedTableColIndex < 0) selectedTableColIndex = selectedGridCol;
            int insertBeforeIndex = selectedTableColIndex + 1;
            int selectedDisplayIndex = dgv.Columns[selectedGridCol].DisplayIndex;

            string input = Microsoft.VisualBasic.Interaction.InputBox("How many columns to insert?", "Insert Columns", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;
            int dataColCount = table.Columns.Count - (table.Columns.Contains(OriginalIndexColumnName) ? 1 : 0);
            var newColNames = new List<string>();
            for (int i = 0; i < count; i++)
            {
                string colName = "Col" + (dataColCount + i);
                while (table.Columns.Contains(colName)) colName += "_";
                var col = table.Columns.Add(colName);
                col.SetOrdinal(insertBeforeIndex + i);
                newColNames.Add(colName);
            }
            for (int r = 0; r < table.Rows.Count; r++)
            {
                foreach (string colName in newColNames)
                    table.Rows[r][colName] = "";
            }
            dgv.Refresh();
            for (int i = 0; i < newColNames.Count; i++)
            {
                if (dgv.Columns.Contains(newColNames[i]))
                    dgv.Columns[newColNames[i]].DisplayIndex = selectedDisplayIndex + 1 + i;
            }
            MarkCurrentTabFileDirty();
            string? fpInsert = GetFilePathForGrid(dgv);
            if (fpInsert != null) PushStructuralUndo(fpInsert, "REMOVE_COLUMNS:" + string.Join("\x01", newColNames), "Insert columns");
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
            UpdateColumnHeaders();
            RefreshColumnHeaders();
            ReapplyAllGridStylesDeferred(dgv);
        }

        private void ColumnOptions_Hide()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null) return;
            var colIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.ColumnIndex)
                .Distinct()
                .Where(c => dgv.Columns[c].Name != OriginalIndexColumnName)
                .ToList();
            if (colIndices.Count == 0)
            {
                MessageBox.Show("Select one or more columns (or cells in those columns) first.", "Column Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            foreach (int c in colIndices)
                dgv.Columns[c].Visible = false;
        }

        private void ColumnOptions_Unhide()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null) return;
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                if (col.Name == OriginalIndexColumnName) continue;
                if (!col.Visible)
                    col.Visible = true;
            }
        }

        private void ColumnOptions_Clone()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            var colIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.ColumnIndex)
                .Distinct()
                .Where(c => dgv.Columns[c].Name != OriginalIndexColumnName)
                .OrderBy(x => x)
                .ToList();
            if (colIndices.Count == 0)
            {
                MessageBox.Show("Select one or more columns (or cells in those columns) first.", "Column Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string input = Microsoft.VisualBasic.Interaction.InputBox("How many copies?", "Clone Columns", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;
            int insertBeforeIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (insertBeforeIndex < 0) insertBeforeIndex = table.Columns.Count;
            int dataColCount = table.Columns.Count - (table.Columns.Contains(OriginalIndexColumnName) ? 1 : 0);
            var newColNames = new List<string>();
            int pos = insertBeforeIndex;
            for (int copy = 0; copy < count; copy++)
            {
                foreach (int srcColIndex in colIndices)
                {
                    string srcColName = table.Columns[srcColIndex].ColumnName;
                    string colName = srcColName + "_copy" + (copy + 1);
                    while (table.Columns.Contains(colName)) colName += "_";
                    var col = table.Columns.Add(colName);
                    col.SetOrdinal(pos);
                    pos++;
                    newColNames.Add(colName);
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        object val = table.Rows[r][srcColName];
                        table.Rows[r][colName] = val == DBNull.Value ? "" : val;
                    }
                }
            }
            MarkCurrentTabFileDirty();
            string? fpClone = GetFilePathForGrid(dgv);
            if (fpClone != null && newColNames.Count > 0) PushStructuralUndo(fpClone, "REMOVE_COLUMNS:" + string.Join("\x01", newColNames), "Clone columns");
            if (newColNames.Count > 0)
            {
                int firstNewColIndex = table.Columns.IndexOf(newColNames[0]);
                dgv.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(firstNewColIndex, dgv.ColumnCount - 1));
            }
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
            UpdateColumnHeaders();
            RefreshColumnHeaders();
            ReapplyAllGridStylesDeferred(dgv);
        }

        private static string EscapeStructuralPayload(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Replace("\x01", "\uE000").Replace("\x02", "\uE001");
        }

        private static string UnescapeStructuralPayload(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Replace("\uE000", "\x01").Replace("\uE001", "\x02");
        }

        private void ColumnOptions_Remove()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            var colIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.ColumnIndex)
                .Distinct()
                .Where(c => dgv.Columns[c].Name != OriginalIndexColumnName)
                .OrderBy(x => x)
                .ToList();
            if (colIndices.Count == 0)
            {
                MessageBox.Show("Select one or more columns (or cells in those columns) first.", "Column Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string? fp = GetFilePathForGrid(dgv);
            if (fp != null)
            {
                var colNames = colIndices.Select(ci => table.Columns[ci].ColumnName).ToList();
                var sb = new System.Text.StringBuilder();
                sb.Append(string.Join("\x01", colIndices.Zip(colNames, (idx, name) => idx + ":" + name)));
                sb.Append("\x02");
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < colIndices.Count; c++)
                    {
                        if (c > 0) sb.Append("\x01");
                        object v = table.Rows[r][colIndices[c]];
                        sb.Append(EscapeStructuralPayload(v == null || v == DBNull.Value ? "" : v.ToString() ?? ""));
                    }
                    sb.Append("\x02");
                }
                PushStructuralUndo(fp, "RESTORE_COLUMNS:" + sb.ToString(), "Remove columns");
            }
            foreach (int ci in colIndices.OrderByDescending(x => x))
            {
                string colName = table.Columns[ci].ColumnName;
                table.Columns.Remove(colName);
            }
            MarkCurrentTabFileDirty();
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
            UpdateColumnHeaders();
            RefreshColumnHeaders();
        }

        private void RowOptions_Add()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            string input = Microsoft.VisualBasic.Interaction.InputBox("How many rows to add?", "Add Rows", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;
            int origColIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (origColIndex < 0) return;
            int startRowIndex = table.Rows.Count;
            int nextOrigIndex = startRowIndex - 1;
            for (int i = 0; i < count; i++)
            {
                var row = table.NewRow();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    if (c == origColIndex)
                        row[c] = nextOrigIndex + i;
                    else
                        row[c] = "";
                }
                table.Rows.Add(row);
            }
            MarkCurrentTabFileDirty();
            string? fpAdd = GetFilePathForGrid(dgv);
            if (fpAdd != null) PushStructuralUndo(fpAdd, $"REMOVE_ROWS:{startRowIndex}:{count}", "Add rows");
            dgv.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(startRowIndex, dgv.RowCount - 1));
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
        }

        private void RowOptions_Insert()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            int selectedGridRow = -1;
            if (dgv.SelectedCells.Count > 0)
            {
                selectedGridRow = dgv.SelectedCells.Cast<DataGridViewCell>()
                    .Where(c => c.RowIndex >= 0)
                    .Select(c => c.RowIndex)
                    .DefaultIfEmpty(-1)
                    .Max();
            }
            if (selectedGridRow < 0 && dgv.CurrentCell != null && dgv.CurrentCell.RowIndex >= 0)
                selectedGridRow = dgv.CurrentCell.RowIndex;
            if (selectedGridRow < 0)
            {
                MessageBox.Show("Select a row (or a cell in the row) after which to insert.", "Insert Rows", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int selectedTableRowIndex = selectedGridRow;
            if (selectedTableRowIndex >= table.Rows.Count) return;
            int origColIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (origColIndex < 0) return;
            int selectedOrigIndex = Convert.ToInt32(table.Rows[selectedTableRowIndex][origColIndex]);
            if (selectedOrigIndex < 0) selectedOrigIndex = -1;

            string input = Microsoft.VisualBasic.Interaction.InputBox("How many rows to insert?", "Insert Rows", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;

            var newRows = new List<DataRow>();
            for (int i = 0; i < count; i++)
            {
                var row = table.NewRow();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    if (c == origColIndex)
                        row[c] = selectedOrigIndex + 1 + i;
                    else
                        row[c] = "";
                }
                newRows.Add(row);
            }
            for (int i = count - 1; i >= 0; i--)
                table.Rows.InsertAt(newRows[i], selectedTableRowIndex + 1);
            for (int r = selectedTableRowIndex + 1 + count; r < table.Rows.Count; r++)
            {
                object v = table.Rows[r][origColIndex];
                if (v != DBNull.Value && v != null)
                    table.Rows[r][origColIndex] = Convert.ToInt32(v) + count;
            }
            MarkCurrentTabFileDirty();
            string? fpInsertRow = GetFilePathForGrid(dgv);
            if (fpInsertRow != null) PushStructuralUndo(fpInsertRow, $"REMOVE_ROWS:{selectedTableRowIndex + 1}:{count}", "Insert rows");
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
        }

        private void RowOptions_Hide()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null) return;
            var rowIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .Where(r => r >= 0)
                .ToList();
            if (rowIndices.Count == 0)
            {
                MessageBox.Show("Select one or more rows (or cells in those rows) first.", "Row Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int currentRow = dgv.CurrentCell?.RowIndex ?? -1;
            if (currentRow >= 0 && rowIndices.Contains(currentRow))
            {
                int targetRow = -1;
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    if (!rowIndices.Contains(i) && !dgv.Rows[i].IsNewRow && dgv.Rows[i].Visible)
                    {
                        targetRow = i;
                        break;
                    }
                }
                if (targetRow >= 0 && dgv.Columns.Count > 0)
                {
                    int col = Math.Max(0, dgv.CurrentCell?.ColumnIndex ?? 0);
                    dgv.CurrentCell = dgv[col, targetRow];
                }
                else
                    dgv.CurrentCell = null;
            }
            foreach (int r in rowIndices)
            {
                if (r < dgv.Rows.Count && !dgv.Rows[r].IsNewRow)
                    dgv.Rows[r].Visible = false;
            }
        }

        private void RowOptions_Unhide()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null) return;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                if (!row.Visible)
                    row.Visible = true;
            }
        }

        private void RowOptions_Clone()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            var rowIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .Where(r => r >= 0 && r < table.Rows.Count)
                .OrderBy(x => x)
                .ToList();
            if (rowIndices.Count == 0)
            {
                MessageBox.Show("Select one or more rows (or cells in those rows) first.", "Row Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string input = Microsoft.VisualBasic.Interaction.InputBox("How many copies?", "Clone Rows", "1");
            if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input.Trim(), out int count) || count <= 0) return;
            int origColIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            if (origColIndex < 0) return;
            int startRowIndex = table.Rows.Count;
            int nextOrigIndex = startRowIndex - 1;
            for (int copy = 0; copy < count; copy++)
            {
                foreach (int srcRowIndex in rowIndices)
                {
                    var newRow = table.NewRow();
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        if (c == origColIndex)
                            newRow[c] = nextOrigIndex++;
                        else
                        {
                            object val = table.Rows[srcRowIndex][c];
                            newRow[c] = val == DBNull.Value ? "" : val;
                        }
                    }
                    table.Rows.Add(newRow);
                }
            }
            MarkCurrentTabFileDirty();
            int addedCount = count * rowIndices.Count;
            string? fpCloneRow = GetFilePathForGrid(dgv);
            if (fpCloneRow != null && addedCount > 0) PushStructuralUndo(fpCloneRow, $"REMOVE_ROWS:{startRowIndex}:{addedCount}", "Clone rows");
            dgv.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(startRowIndex, dgv.RowCount - 1));
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
        }

        private void RowOptions_Remove()
        {
            var dgv = GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable table) return;
            var rowIndices = dgv.SelectedCells.Cast<DataGridViewCell>()
                .Select(c => c.RowIndex)
                .Distinct()
                .Where(r => r >= 0 && r < table.Rows.Count)
                .OrderBy(x => x)
                .ToList();
            if (rowIndices.Count == 0)
            {
                MessageBox.Show("Select one or more rows (or cells in those rows) first.", "Row Options", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string? fp = GetFilePathForGrid(dgv);
            if (fp != null)
            {
                var sb = new System.Text.StringBuilder();
                sb.Append(rowIndices.Count);
                sb.Append("\x02");
                sb.Append(string.Join("\x01", rowIndices));
                sb.Append("\x02");
                int colCount = table.Columns.Count;
                foreach (int ri in rowIndices)
                {
                    var row = table.Rows[ri];
                    for (int c = 0; c < colCount; c++)
                    {
                        if (c > 0) sb.Append("\x01");
                        object v = row[c];
                        sb.Append(EscapeStructuralPayload(v == null || v == DBNull.Value ? "" : v.ToString() ?? ""));
                    }
                    sb.Append("\x02");
                }
                PushStructuralUndo(fp, "RESTORE_ROWS:" + sb.ToString(), "Remove rows");
            }
            foreach (int ri in rowIndices.OrderByDescending(x => x))
                table.Rows.RemoveAt(ri);
            MarkCurrentTabFileDirty();
            ApplyRowStyles(dgv);
            ApplyHeaderCellColors(dgv);
        }

        private void ContextMenuStrip1_Opening(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            contextMenuStrip1.Items.Clear();
            _contextMenuRestoreMode = 0;
            var dgv = contextMenuStrip1.SourceControl as DataGridView ?? GetCurrentGrid();
            if (dgv == null || dgv.DataSource is not DataTable) return;
            var pt = dgv.PointToClient(Cursor.Position);
            var hit = dgv.HitTest(pt.X, pt.Y);
            bool menuFromHit = false;
            if (hit.Type == DataGridViewHitTestType.ColumnHeader && hit.ColumnIndex >= 0 && dgv.Columns[hit.ColumnIndex].Name != OriginalIndexColumnName)
            {
                int colIdx = hit.ColumnIndex;
                dgv.BeginInvoke(new Action(() =>
                {
                    if (colIdx < dgv.Columns.Count && dgv.Columns[colIdx].Name != OriginalIndexColumnName)
                    {
                        dgv.CurrentCell = dgv.Rows[0].Cells[colIdx];
                        dgv.Columns[colIdx].Selected = true;
                    }
                }));
                menuFromHit = true;
            }
            else if (hit.Type == DataGridViewHitTestType.RowHeader && hit.RowIndex >= 0 && hit.RowIndex < dgv.Rows.Count && !dgv.Rows[hit.RowIndex].IsNewRow)
            {
                int rowIdx = hit.RowIndex;
                dgv.BeginInvoke(new Action(() =>
                {
                    if (rowIdx >= 0 && rowIdx < dgv.Rows.Count && !dgv.Rows[rowIdx].IsNewRow)
                    {
                        dgv.CurrentCell = dgv.Rows[rowIdx].Cells[0];
                        dgv.Rows[rowIdx].Selected = true;
                    }
                }));
                menuFromHit = true;
            }
            var cells = dgv.SelectedCells.Cast<DataGridViewCell>().ToList();
            if (cells.Count == 0 && !menuFromHit) return;
            int dataRowCount = dgv.Rows.Count;
            int dataColCount = dgv.Columns.Cast<DataGridViewColumn>().Count(c => c.Name != OriginalIndexColumnName);
            bool fullColumn;
            bool fullRow;
            if (menuFromHit)
            {
                fullColumn = hit.Type == DataGridViewHitTestType.ColumnHeader;
                fullRow = hit.Type == DataGridViewHitTestType.RowHeader;
            }
            else
            {
                var selectedCols = cells.Select(c => c.ColumnIndex).Distinct().Where(c => dgv.Columns[c].Name != OriginalIndexColumnName).ToList();
                var selectedRows = cells.Select(c => c.RowIndex).Distinct().Where(r => r >= 0 && r < dataRowCount).ToList();
                fullColumn = selectedCols.Count > 0 && selectedCols.All(col =>
                    cells.Count(c => c.ColumnIndex == col) == dataRowCount &&
                    cells.Where(c => c.ColumnIndex == col).Select(c => c.RowIndex).Distinct().Count() == dataRowCount);
                fullRow = selectedRows.Count > 0 && selectedRows.All(row =>
                {
                    var rowDataCells = cells.Where(c => c.RowIndex == row && dgv.Columns[c.ColumnIndex].Name != OriginalIndexColumnName).ToList();
                    return rowDataCells.Count == dataColCount && rowDataCells.Select(c => c.ColumnIndex).Distinct().Count() == dataColCount;
                });
                if (!fullColumn && !fullRow)
                {
                    if (hit.Type == DataGridViewHitTestType.ColumnHeader && hit.ColumnIndex >= 0 && dgv.Columns[hit.ColumnIndex].Name != OriginalIndexColumnName)
                    {
                        int colIdx = hit.ColumnIndex;
                        dgv.BeginInvoke(new Action(() =>
                        {
                            if (colIdx < dgv.Columns.Count && dgv.Columns[colIdx].Name != OriginalIndexColumnName)
                            {
                                dgv.CurrentCell = dgv.Rows[0].Cells[colIdx];
                                dgv.Columns[colIdx].Selected = true;
                            }
                        }));
                        fullColumn = true;
                        fullRow = false;
                    }
                    else if (hit.Type == DataGridViewHitTestType.RowHeader && hit.RowIndex >= 0 && hit.RowIndex < dgv.Rows.Count && !dgv.Rows[hit.RowIndex].IsNewRow)
                    {
                        int rowIdx = hit.RowIndex;
                        dgv.BeginInvoke(new Action(() =>
                        {
                            if (rowIdx >= 0 && rowIdx < dgv.Rows.Count && !dgv.Rows[rowIdx].IsNewRow)
                            {
                                dgv.CurrentCell = dgv.Rows[rowIdx].Cells[0];
                                dgv.Rows[rowIdx].Selected = true;
                            }
                        }));
                        fullColumn = false;
                        fullRow = true;
                    }
                }
            }
            if (fullColumn && !fullRow)
            {
                _contextMenuRestoreMode = 1;
                contextMenuStrip1.Items.Add(addColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(insertColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(hideColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(removeColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(unhideColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(cloneColumnsToolStripMenuItem);
            }
            else if (fullRow && !fullColumn)
            {
                _contextMenuRestoreMode = 2;
                contextMenuStrip1.Items.Add(addRowsToolStripMenuItem);
                contextMenuStrip1.Items.Add(insertRowsToolStripMenuItem);
                contextMenuStrip1.Items.Add(hideRowsToolStripMenuItem);
                contextMenuStrip1.Items.Add(removeRowsToolStripMenuItem);
                contextMenuStrip1.Items.Add(unhideRowsToolStripMenuItem);
                contextMenuStrip1.Items.Add(cloneRowsToolStripMenuItem);
            }
            else if (fullColumn && fullRow)
            {
                _contextMenuRestoreMode = 1;
                contextMenuStrip1.Items.Add(addColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(insertColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(hideColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(removeColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(unhideColumnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(cloneColumnsToolStripMenuItem);
            }
            if (contextMenuStrip1.Items.Count == 0)
                e.Cancel = true;
        }

        private void ContextMenuStrip1_Closed(object? sender, ToolStripDropDownClosedEventArgs e)
        {
            if (_contextMenuRestoreMode == 1)
            {
                columnsToolStripMenuItem.DropDownItems.Add(addColumnsToolStripMenuItem);
                columnsToolStripMenuItem.DropDownItems.Add(insertColumnsToolStripMenuItem);
                columnsToolStripMenuItem.DropDownItems.Add(hideColumnsToolStripMenuItem);
                columnsToolStripMenuItem.DropDownItems.Add(removeColumnsToolStripMenuItem);
                columnsToolStripMenuItem.DropDownItems.Add(unhideColumnsToolStripMenuItem);
                columnsToolStripMenuItem.DropDownItems.Add(cloneColumnsToolStripMenuItem);
            }
            else if (_contextMenuRestoreMode == 2)
            {
                rowsToolStripMenuItem.DropDownItems.Add(addRowsToolStripMenuItem);
                rowsToolStripMenuItem.DropDownItems.Add(insertRowsToolStripMenuItem);
                rowsToolStripMenuItem.DropDownItems.Add(hideRowsToolStripMenuItem);
                rowsToolStripMenuItem.DropDownItems.Add(removeRowsToolStripMenuItem);
                rowsToolStripMenuItem.DropDownItems.Add(unhideRowsToolStripMenuItem);
                rowsToolStripMenuItem.DropDownItems.Add(cloneRowsToolStripMenuItem);
            }
            if (_contextMenuRestoreMode != 0)
            {
                // Re-add parent items so the strip is never empty; otherwise the next right-click won't open the menu.
                contextMenuStrip1.Items.Add(columnsToolStripMenuItem);
                contextMenuStrip1.Items.Add(rowsToolStripMenuItem);
            }
            _contextMenuRestoreMode = 0;
        }

        private void MarkCurrentTabFileDirty()
        {
            string? filePath = (tabControl.SelectedTab?.Tag as TabTag)?.FilePath ?? currentFilePath;
            if (string.IsNullOrEmpty(filePath)) return;
            string pathNorm = NormalizeFilePath(filePath);
            _dirtyFilePaths.Add(pathNorm);
            UpdateDirtyIndicators(filePath);
        }

        private DataGridView CreateDataGridView()
        {
            var dgv = new DataGridView();
            dgv.Dock = DockStyle.Fill;
            dgv.MultiSelect = true;
            dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable;
            dgv.AllowUserToAddRows = false;
            dgv.ColumnHeadersHeight = 24;
            dgv.RowHeadersVisible = true;
            dgv.EnableHeadersVisualStyles = false;
            if (_gridFont != null)
                dgv.Font = (Font)_gridFont.Clone();
            else
                dgv.Font = new Font(SystemFonts.DefaultFont.FontFamily, 9f);
            dgv.CellPainting += DataGridView_CellPainting;
            dgv.CellValueChanged += DataGridView_CellValueChanged;
            dgv.CellBeginEdit += DataGridView_CellBeginEdit;
            dgv.CellEndEdit += DataGridView_CellEndEdit;
            dgv.KeyDown += DataGridView_KeyDown;
            dgv.ColumnHeaderMouseClick += DataGridView_ColumnHeaderMouseClick;
            dgv.RowHeaderMouseClick += DataGridView_RowHeaderMouseClick;
            dgv.ColumnWidthChanged += DataGridView_ColumnWidthChanged;
            dgv.RowHeightChanged += DataGridView_RowHeightChanged;
            dgv.MouseMove += DataGuideHeaderMouseMoveFromTabOrGrid;
            dgv.MouseLeave += HeaderTooltipGrid_MouseLeave;
            dgv.ContextMenuStrip = contextMenuStrip1;
            SetDoubleBuffered(dgv);
            ApplyGridColorsToGrid(dgv);
            return dgv;
        }

        private async void LoadFile(string filePath, Color? workspaceColor = null, string? workspaceFolderPath = null, Action? onComplete = null)
        {
            string pathNorm = NormalizeFilePath(filePath);

            foreach (TabPage tab in tabControl.TabPages)
            {
                string tabPath = NormalizeFilePath((tab.Tag as TabTag)?.FilePath ?? tab.Tag?.ToString());
                if (string.Equals(tabPath, pathNorm, StringComparison.OrdinalIgnoreCase))
                {
                    tabControl.SelectedTab = tab;
                    onComplete?.Invoke();
                    return;
                }
            }

            bool showProgress =
                _loadProgressPanel != null &&
                _loadProgressBar != null &&
                _loadProgressLabel != null;

            if (showProgress)
            {
                _loadProgressPanel!.Visible = true;
                _loadProgressBar!.Value = 0;
                _loadProgressLabel!.Text = "Loading...";
                _loadProgressPanel.Refresh();
            }

            long fileLen = new FileInfo(filePath).Length;

            IProgress<double>? progress = null;
            if (showProgress)
            {
                progress = new Progress<double>(p =>
                {
                    int percent = (int)Math.Round(p);
                    percent = Math.Clamp(percent, 0, 100);
                    _loadProgressBar!.Value = percent;

                    if (percent < 50)
                        _loadProgressLabel!.Text = $"Loading {percent * fileLen / (50 * 1024)} / {fileLen / 1024} KB";
                    else
                        _loadProgressLabel!.Text = "Finalizing table...";
                });
            }

            DataTable dataTable;
            try
            {
                dataTable = await Task.Run(() =>
                    ReadFileToDataTable(filePath, fileLen, progress));
            }
            catch
            {
                if (showProgress)
                    _loadProgressPanel!.Visible = false;

                onComplete?.Invoke();
                return;
            }

            // ---------------- FINALIZATION PHASE ----------------
            // Populate UI + DataGridView on UI thread
            var tabTag = new TabTag
            {
                FilePath = filePath,
                WorkspaceFolderPath = workspaceFolderPath,
                WorkspaceBackColor = workspaceColor
            };
            TabPage tabPage = new TabPage(Path.GetFileNameWithoutExtension(filePath) + GetTabWidthPaddingString())
            {
                Tag = tabTag
            };

            if (workspaceColor.HasValue)
                tabPage.BackColor = workspaceColor.Value;

            DataGridView dgv = CreateDataGridView();
            tabPage.Controls.Add(dgv);
            tabPage.MouseMove += DataGuideHeaderMouseMoveFromTabOrGrid;
            tabControl.TabPages.Add(tabPage);
            tabControl.SelectedTab = tabPage;

            // Bind the full table at once so row 0 (header) is never lost (avoids CopyToDataTable/ImportRow issues)
            dgv.DataSource = dataTable;
            if (progress != null)
            {
                _loadProgressBar!.Value = 75;
                _loadProgressLabel!.Text = "Finalizing table...";
                await Task.Yield();
            }

            dataGridView = dgv;

            HideOriginalIndexColumn();
            ApplyDefaultHeaderLocksIfEnabled();
            UpdateColumnHeaders();
            ApplyFrozenColumns();
            ApplyFrozenRows();
            ApplyRowStyles();
            DisableColumnSorting();
            if (_adaptiveColumnWidthMenuItem?.Checked == true)
            {
                var widths = new int[dgv.Columns.Count];
                for (int i = 0; i < dgv.Columns.Count; i++)
                    widths[i] = dgv.Columns[i].Width;
                _savedColumnWidthsByTab[tabPage] = widths;
                _inGroupedColumnWidthSync = true;
                try { dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells); }
                finally { _inGroupedColumnWidthSync = false; }
            }
            string pathNormalize = NormalizeFilePath(filePath);
            if (!string.IsNullOrEmpty(pathNorm))
            {
                if (_savedColumnWidthsByFile.TryGetValue(pathNorm, out int[]? widths))
                {
                    _inGroupedColumnWidthSync = true;
                    try
                    {
                        for (int i = 0; i < dgv.Columns.Count && i < widths.Length; i++)
                            dgv.Columns[i].Width = widths[i];
                    }
                    finally { _inGroupedColumnWidthSync = false; }
                }
                if (_savedRowHeightsByFile.TryGetValue(pathNorm, out int[]? heights))
                {
                    _inGroupedRowHeightSync = true;
                    try
                    {
                        for (int i = 0; i < dgv.Rows.Count && i < heights.Length; i++)
                        {
                            if (dgv.Rows[i].IsNewRow) continue;
                            dgv.Rows[i].Height = heights[i];
                        }
                    }
                    finally { _inGroupedRowHeightSync = false; }
                }
            }
            if (dgv.DataSource is DataTable initialTable)
                StoreSavedContent(filePath, GetContentStringFromTable(initialTable));
            UpdateDirtyIndicators(filePath);

            if (showProgress)
            {
                _loadProgressBar!.Value = 100;
                _loadProgressLabel!.Text = "Done";
                _loadProgressPanel.Refresh();
                await Task.Delay(50);
                _loadProgressPanel.Visible = false;
                _loadProgressBar!.Value = 0;
            }

            onComplete?.Invoke();
        }

        private static DataTable ReadFileToDataTable(string filePath, long fileLen, IProgress<double>? progress)
        {
            var dataTable = new DataTable();
            Stream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, ProgressStreamReaderBufferSize, FileOptions.SequentialScan);

            using (stream)
            using (var reader = new StreamReader(stream, Encoding.UTF8, true, ProgressStreamReaderBufferSize))
            {
                string? line = reader.ReadLine();
                if (line == null)
                    return dataTable;

                string[] headerValues = line.Split('\t');

                for (int i = 0; i < headerValues.Length; i++)
                {
                    string name = headerValues[i];
                    if (string.IsNullOrWhiteSpace(name)) name = "Col" + i;
                    else name = name.Trim();
                    if (dataTable.Columns.Contains(name))
                        name = name + "_" + i;
                    dataTable.Columns.Add(name);
                }
                dataTable.Columns.Add(OriginalIndexColumnName, typeof(int));
                dataTable.ExtendedProperties["HeaderRow"] = headerValues;

                dataTable.BeginLoadData();

                // Row 0 = editable header row (column text)
                object[] headerRowObj = new object[dataTable.Columns.Count];
                for (int c = 0; c < dataTable.Columns.Count - 1; c++)
                    headerRowObj[c] = c < headerValues.Length ? headerValues[c] : "";
                headerRowObj[dataTable.Columns.Count - 1] = -1;
                dataTable.Rows.Add(headerRowObj);

                int rowIndex = 0;
                long lastReportedBytes = 0;
                const long MinBytesDelta = 128 * 1024; // report every 128 KB

                while ((line = reader.ReadLine()) != null)
                {
                    string[] rowValues = line.Split('\t');
                    object[] values = new object[dataTable.Columns.Count];

                    for (int c = 0; c < dataTable.Columns.Count - 1; c++)
                        values[c] = c < rowValues.Length ? rowValues[c] : "";

                    values[dataTable.Columns.Count - 1] = rowIndex++;
                    dataTable.Rows.Add(values);

                    // 🔹 Report progress based on bytes read
                    if (progress != null && stream.CanSeek)
                    {
                        long bytesRead = stream.Position;
                        if (bytesRead - lastReportedBytes >= MinBytesDelta)
                        {
                            lastReportedBytes = bytesRead;
                            double percent = (bytesRead / (double)fileLen) * 50.0; // first 50% = parsing
                            progress.Report(percent);
                        }
                    }
                }

                dataTable.EndLoadData();
            }

            // 🔹 After parsing, report 50% (parsing done)
            progress?.Report(50.0);

            return dataTable;
        }

        private sealed class ProgressReportingStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _totalBytes;
            private readonly Action<long, long> _report;

            public ProgressReportingStream(Stream inner, long totalBytes, Action<long, long> report)
            {
                _inner = inner;
                _totalBytes = totalBytes;
                _report = report;
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int n = _inner.Read(buffer, offset, count);
                if (n > 0) _report(_inner.Position, _totalBytes);
                return n;
            }

            public override long Position { get => _inner.Position; set => _inner.Position = value; }
            public override long Length => _inner.Length;
            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => _inner.CanSeek;
            public override bool CanWrite => false;
            public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override void Flush() => _inner.Flush();
            protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
        }

        private void HideOriginalIndexColumn()
        {
            if (dataGridView == null) return;
            if (dataGridView.Columns.Contains(OriginalIndexColumnName))
                dataGridView.Columns[OriginalIndexColumnName].Visible = false;
        }

        // Extra width (px) to the right of each tab (smaller = narrower tabs, less gap between tabs)
        private const int TabRightPaddingPixels = 40;
        // Space reserved for the close (X) button on each tab
        private const int TabCloseButtonWidth = 15;
        // Active (selected) tab background color — change these to adjust the red
        private static readonly Color ActiveTabBackColorDark = Color.FromArgb(149, 215, 251);
        private static readonly Color ActiveTabBackColorLight = Color.FromArgb(255, 200, 200);

        private string GetTabWidthPaddingString()
        {
            using (var g = tabControl.CreateGraphics())
            {
                var spaceSize = TextRenderer.MeasureText(g, " ", tabControl.Font);
                if (spaceSize.Width <= 0) return "    ";
                int totalPadding = TabRightPaddingPixels + TabCloseButtonWidth;
                int count = (int)Math.Ceiling((double)totalPadding / spaceSize.Width);
                return new string(' ', Math.Max(2, count + 1));
            }
        }

        private string GetTabDisplayText(TabPage page)
        {
            if (page.Tag is TabTag tag && !string.IsNullOrEmpty(tag.FilePath))
            {
                string name = Path.GetFileNameWithoutExtension(tag.FilePath);
                if (_dirtyFilePaths.Contains(NormalizeFilePath(tag.FilePath)))
                    name += "*";
                return name;
            }
            return page.Text.TrimEnd();
        }

        private void DataGridView_CellBeginEdit(object? sender, DataGridViewCellCancelEventArgs e)
        {
            var dgv = sender as DataGridView;
            if (dgv == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (dgv.Columns[e.ColumnIndex].Name == OriginalIndexColumnName) return;
            if (e.RowIndex >= dgv.Rows.Count || dgv.Rows[e.RowIndex].IsNewRow) return;
            string currentVal = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? "";
            _pendingUndoCell = (e.RowIndex, e.ColumnIndex, currentVal);
        }

        private void DataGridView_CellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (_isApplyingUndo) return;
            var dgv = sender as DataGridView;
            if (dgv == null || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (dgv.Columns[e.ColumnIndex].Name == OriginalIndexColumnName) return;
            if (!_pendingUndoCell.HasValue) return;
            var (row, col, value) = _pendingUndoCell.Value;
            if (row != e.RowIndex || col != e.ColumnIndex) return;
            _pendingUndoCell = null;
            if (tabControl.SelectedTab?.Tag is not TabTag tag || string.IsNullOrEmpty(tag.FilePath)) return;
            string columnName = (dgv.Columns[col].HeaderText ?? dgv.Columns[col].Name ?? "").Trim();
            if (string.IsNullOrEmpty(columnName)) columnName = "Column" + col;
            string rowName = row == 0 ? "Header" : (dgv.Rows[row].HeaderCell.Value?.ToString()?.Trim() ?? "");
            if (string.IsNullOrEmpty(rowName))
            {
                int firstCol = -1;
                for (int c = 0; c < dgv.Columns.Count; c++)
                {
                    if (dgv.Columns[c].Name != OriginalIndexColumnName) { firstCol = c; break; }
                }
                if (firstCol >= 0 && firstCol < dgv.Rows[row].Cells.Count)
                    rowName = dgv.Rows[row].Cells[firstCol].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(rowName)) rowName = "Row " + (row + 1);
            }
            string newVal = dgv.Rows[row].Cells[col].Value?.ToString() ?? "";
            string oldVal = value ?? "";
            string description = $"{columnName} of {rowName} edited from {oldVal} to {newVal}";
            PushUndoEntry(tag.FilePath, new List<(int row, int col, string value)> { (row, col, value) }, description,
                columnName, rowName, oldVal, newVal);
        }

        private void DataGridView_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView == null || tabControl.SelectedTab?.Tag is not TabTag tag || string.IsNullOrEmpty(tag.FilePath)) return;
            if (e.RowIndex == 0)
            {
                SyncHeaderRowFromRow0();
                UpdateColumnHeaders();
            }
            int firstDataCol = GetFirstDataColumnIndex(dataGridView);
            if (e.RowIndex >= 0 && firstDataCol >= 0 && e.ColumnIndex == firstDataCol)
            {
                ApplyRowStyles();
                ApplyHeaderCellColors(dataGridView);
            }
            string pathNorm = NormalizeFilePath(tag.FilePath);
            _dirtyFilePaths.Add(pathNorm);
            UpdateDirtyIndicators(tag.FilePath);
            if (dataGridView.DataSource is DataTable table)
            {
                string current = GetContentStringFromTable(table);
                if (_savedContentByFilePath.TryGetValue(pathNorm, out string? saved) && current == saved)
                {
                    _dirtyFilePaths.Remove(pathNorm);
                    UpdateDirtyIndicators(tag.FilePath);
                }
            }
        }

        private static string GetContentStringFromTable(DataTable table)
        {
            var sb = new StringBuilder();
            int origColIndex = table.Columns.IndexOf(OriginalIndexColumnName);
            for (int r = 0; r < table.Rows.Count; r++)
            {
                var values = new List<string>();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    if (c == origColIndex) continue;
                    values.Add(table.Rows[r][c]?.ToString() ?? "");
                }
                sb.AppendLine(string.Join("\t", values));
            }
            return sb.ToString();
        }

        private void StoreSavedContent(string filePath, string content)
        {
            string pathNorm = NormalizeFilePath(filePath);
            if (string.IsNullOrEmpty(pathNorm)) return;
            _savedContentByFilePath[pathNorm] = content;
        }

        private string? GetSavedContent(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            return _savedContentByFilePath.TryGetValue(pathNorm, out string? content) ? content : null;
        }

        private void UpdateDirtyIndicators(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            string displayName = Path.GetFileNameWithoutExtension(filePath);
            bool dirty = _dirtyFilePaths.Contains(pathNorm);

            foreach (TabPage tab in tabControl.TabPages)
            {
                if (tab.Tag is TabTag t && string.Equals(NormalizeFilePath(t.FilePath), pathNorm, StringComparison.OrdinalIgnoreCase))
                {
                    tab.Text = displayName + (dirty ? "*" : "") + GetTabWidthPaddingString();
                    break;
                }
            }

            TreeNode? node = FindTreeNodeByFilePath(filePath);
            if (node != null)
                node.Text = Path.GetFileName(filePath) + (dirty ? "*" : "");
        }

        private TreeNode? FindTreeNodeByFilePath(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            foreach (TreeNode root in treeView.Nodes)
            {
                foreach (TreeNode child in root.Nodes)
                {
                    if (child.Tag is string tagPath && string.Equals(NormalizeFilePath(tagPath), pathNorm, StringComparison.OrdinalIgnoreCase))
                        return child;
                }
            }
            return null;
        }

        private void SelectWorkspaceNodeForFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            TreeNode? node = FindTreeNodeByFilePath(filePath);
            if (node != null)
            {
                // Ensure workspace is expanded and file node is visible and selected.
                if (node.Parent != null)
                    node.Parent.Expand();
                treeView.SelectedNode = node;
                node.EnsureVisible();
            }
            else
            {
                treeView.SelectedNode = null;
            }
        }

        private void CompareWorkspaceFile()
        {
            TreeNode? fileNode = treeView.SelectedNode;
            if (fileNode == null || fileNode.Parent == null || fileNode.Tag is not string filePath)
                return;

            // Collect available workspaces (root nodes with WorkspaceTag)
            var workspaces = treeView.Nodes.Cast<TreeNode>()
                .Where(n => n.Tag is WorkspaceTag)
                .ToList();

            if (workspaces.Count == 0)
            {
                MessageBox.Show(this, "No workspaces are defined.", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Build dialog UI: choose source and target workspace.
            using (var dlg = new Form())
            {
                dlg.Text = "Select source and target workspaces";
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.MinimizeBox = false;
                dlg.MaximizeBox = false;
                dlg.ShowInTaskbar = false;
                dlg.ClientSize = new Size(420, 140);

                var lblSource = new Label { Text = "Source workspace:", AutoSize = true, Location = new Point(12, 15) };
                var cboSource = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new Point(140, 12),
                    Width = 260
                };

                var lblTarget = new Label { Text = "Target workspace:", AutoSize = true, Location = new Point(12, 50) };
                var cboTarget = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Location = new Point(140, 47),
                    Width = 260
                };

                foreach (var ws in workspaces)
                {
                    cboSource.Items.Add(ws.Text);
                    cboTarget.Items.Add(ws.Text);
                }

                // Default source to the file's own workspace (its parent), target to a different workspace if available.
                int sourceIndex = workspaces.FindIndex(ws => ws == fileNode.Parent);
                if (sourceIndex < 0) sourceIndex = 0;
                cboSource.SelectedIndex = sourceIndex;
                cboTarget.SelectedIndex = workspaces.Count > 1 ? (sourceIndex == 0 ? 1 : 0) : sourceIndex;

                var btnOk = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    Location = new Point(dlg.ClientSize.Width - 180, 100),
                    Width = 75
                };
                var btnCancel = new Button
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                    Location = new Point(dlg.ClientSize.Width - 95, 100),
                    Width = 75
                };

                dlg.Controls.Add(lblSource);
                dlg.Controls.Add(cboSource);
                dlg.Controls.Add(lblTarget);
                dlg.Controls.Add(cboTarget);
                dlg.Controls.Add(btnOk);
                dlg.Controls.Add(btnCancel);
                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog(this) != DialogResult.OK)
                    return;

                if (cboSource.SelectedIndex < 0 || cboTarget.SelectedIndex < 0)
                    return;

                var sourceWorkspace = workspaces[cboSource.SelectedIndex];
                var targetWorkspace = workspaces[cboTarget.SelectedIndex];

                if (sourceWorkspace.Tag is not WorkspaceTag sourceTag || targetWorkspace.Tag is not WorkspaceTag targetTag)
                    return;

                string fileName = Path.GetFileName(filePath);

                string sourceFolder = sourceTag.FolderPath;
                string targetFolder = targetTag.FolderPath;
                if (string.IsNullOrWhiteSpace(sourceFolder) || string.IsNullOrWhiteSpace(targetFolder))
                {
                    MessageBox.Show(this, "One or both selected workspaces do not have a valid folder path.", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string sourceFile = Path.Combine(sourceFolder, fileName);
                string targetFile = Path.Combine(targetFolder, fileName);

                if (!File.Exists(sourceFile))
                {
                    MessageBox.Show(this, $"Source file not found:\n{sourceFile}", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!File.Exists(targetFile))
                {
                    MessageBox.Show(this, $"Target file not found:\n{targetFile}", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string exePath = "C:\\Users\\djsch\\Github\\D2Compare-fork\\src\\D2Compare.UI\\bin\\Debug\\net8.0\\D2Compare.exe";

                if (!File.Exists(exePath))
                {
                    MessageBox.Show(this, $"D2Compare.exe was not found in:\n{exePath}", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = exePath,
                        Arguments = $"--source \"{sourceFolder}\" --target \"{targetFolder}\" --file \"{fileName}\"",
                        UseShellExecute = false
                    };
                    Process.Start(startInfo);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Failed to start D2Compare.exe:\n{ex.Message}", "Compare files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void TabControl_PaintTab(object? sender, PaintTabEventArgs e)
        {
            TabPage page = e.Page;
            bool selected = e.Selected;
            Rectangle bounds = e.Bounds;
            Color backColor;
            if (selected)
                backColor = _darkMode ? ActiveTabBackColorDark : ActiveTabBackColorLight;
            else
            {
                backColor = (page.Tag as TabTag)?.WorkspaceBackColor ?? page.BackColor;
                if (backColor == Color.Empty || backColor.A == 0)
                    backColor = _darkMode ? DarkPanel : SystemColors.Control;
            }
            using (var brush = new SolidBrush(backColor))
                e.Graphics.FillRectangle(brush, bounds);
            const int tabTextToCloseGap = 12;
            int textWidth = Math.Max(0, bounds.Width - 4 - TabCloseButtonWidth - tabTextToCloseGap);
            Rectangle textBounds = new Rectangle(bounds.X + 4, bounds.Y + 2, textWidth, bounds.Height - 4);
            string displayText = GetTabDisplayText(page);
            using (var boldFont = new Font(tabControl.Font, FontStyle.Bold))
                TextRenderer.DrawText(e.Graphics, displayText, boldFont, textBounds, TextColorTabActive, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            int xLeft = bounds.Right - TabCloseButtonWidth - 2;
            int xTop = bounds.Y + (bounds.Height - TabCloseButtonWidth) / 2;
            Rectangle xRect = new Rectangle(xLeft, xTop, TabCloseButtonWidth, TabCloseButtonWidth);
            using (var redBrush = new SolidBrush(Color.FromArgb(255, 180, 180)))
                e.Graphics.FillRectangle(redBrush, xRect);
            int pad = 3;
            int x1 = xRect.X + pad;
            int x2 = xRect.Right - pad;
            int y1 = xRect.Y + pad;
            int y2 = xRect.Bottom - pad;
            using (var blackPen = new Pen(Color.Black, 1.5f))
            {
                e.Graphics.DrawLine(blackPen, x1, y1, x2, y2);
                e.Graphics.DrawLine(blackPen, x2, y1, x1, y2);
            }
        }

        private void TabControl_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            for (int i = 0; i < tabControl.TabCount; i++)
            {
                Rectangle r = tabControl.GetTabRect(i);
                const int tabCloseLeftOffset = 2;
                if (r.Contains(e.Location) && e.X >= r.Right - TabCloseButtonWidth - tabCloseLeftOffset)
                {
                    TabPage tab = tabControl.TabPages[i];
                    tabControl.TabPages.RemoveAt(i);
                    tab.Dispose();
                    return;
                }
            }
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl.SelectedTab != null && tabControl.SelectedTab.Tag != null)
            {
                string filePath = (tabControl.SelectedTab.Tag as TabTag)?.FilePath ?? tabControl.SelectedTab.Tag.ToString() ?? "";
                currentFilePath = filePath;
                dataGridView = GetGridForTab(tabControl.SelectedTab);
                // Sync workspace tree selection with the active tab's file, if present.
                SelectWorkspaceNodeForFile(filePath);
            }
            else
            {
                dataGridView = null;
            }
            RefreshEditHistoryListBox();
            RunSearch();
        }

        private void LoadFileToGrid(string filePath)
        {
            string pathNorm = NormalizeFilePath(filePath);
            if (fileCache.TryGetValue(pathNorm, out DataTable cachedData))
            {
                UpdateDataGridView(cachedData);
                return;
            }

            DataTable dataTable = new DataTable();

            using (var reader = new StreamReader(filePath))
            {
                string line;

                if ((line = reader.ReadLine()) != null)
                {
                    string[] headerValues = line.Split('\t');
                    for (int i = 0; i < headerValues.Length; i++)
                        dataTable.Columns.Add("Col" + i);
                    dataTable.Columns.Add(OriginalIndexColumnName, typeof(int));
                    dataTable.ExtendedProperties["HeaderRow"] = headerValues;
                }

                int rowIndex = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] rowValues = line.Split('\t');
                    object[] values = new object[dataTable.Columns.Count];
                    for (int c = 0; c < dataTable.Columns.Count - 1; c++)
                        values[c] = c < rowValues.Length ? rowValues[c] : "";
                    values[dataTable.Columns.Count - 1] = rowIndex++;
                    dataTable.Rows.Add(values);
                }
            }

            fileCache[pathNorm] = dataTable;
            UpdateDataGridView(dataTable);
            StoreSavedContent(filePath, GetContentStringFromTable(dataTable));
            _dirtyFilePaths.Remove(pathNorm);
            UpdateDirtyIndicators(filePath);
        }

        private void UpdateDataGridView(DataTable dataTable)
        {
            dataGridView.Invoke((Action)(() =>
            {
                dataGridView.DataSource = dataTable;
                HideOriginalIndexColumn();
                UpdateColumnHeaders();
                ApplyRowStyles();
                RefreshColumnHeaders();
                DisableColumnSorting();
            }));
        }

        private void ApplyGridColorsToGrid(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = _gridHeaderColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = TextColorHeaderActive;
            dgv.RowHeadersDefaultCellStyle.BackColor = _gridHeaderColor;
            dgv.RowHeadersDefaultCellStyle.ForeColor = TextColorHeaderActive;
            dgv.DefaultCellStyle.ForeColor = TextColorCellActive;
            if (dgv == dataGridView)
                ApplyHeaderCellColors(dgv);
        }

        private void ApplyHeaderCellColors(DataGridView dgv)
        {
            if (dgv == null) return;
            Color headerFore = TextColorHeaderActive;
            int firstDataCol = GetFirstDataColumnIndex(dgv);
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                // Header labels (A,B,C,D) always use header color, never frozen color
                col.HeaderCell.Style.BackColor = _gridHeaderColor;
                col.HeaderCell.Style.ForeColor = headerFore;
            }
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                bool expansionRow = IsExpansionRow(dgv, row.Index, firstDataCol);
                // Header labels (0,1,2,3) always use header/expansion color, never frozen color
                Color back = (row.Index == 0) ? _gridHeaderColor : (expansionRow ? _gridExpansionRowColor : _gridHeaderColor);
                row.HeaderCell.Style.BackColor = back;
                row.HeaderCell.Style.ForeColor = headerFore;
            }
        }

        private void ReapplyAllGridStylesDeferred(DataGridView dgv)
        {
            if (dgv == null || dgv.IsDisposed) return;
            dgv.BeginInvoke(new Action(() =>
            {
                if (dgv.IsDisposed || !dgv.IsHandleCreated || dgv.DataSource == null) return;
                dgv.EnableHeadersVisualStyles = false;
                dgv.ColumnHeadersDefaultCellStyle.BackColor = _gridHeaderColor;
                dgv.ColumnHeadersDefaultCellStyle.ForeColor = TextColorHeaderActive;
                dgv.RowHeadersDefaultCellStyle.BackColor = _gridHeaderColor;
                dgv.RowHeadersDefaultCellStyle.ForeColor = TextColorHeaderActive;
                ApplyRowStyles(dgv);
                ApplyHeaderCellColors(dgv);
                if (dgv == dataGridView)
                {
                    UpdateColumnHeaders();
                    RefreshColumnHeaders();
                }
                dgv.Invalidate();
            }));
        }

        private static Color GetContrastingForeColor(Color back)
        {
            return (back.R * 299 + back.G * 587 + back.B * 114) / 1000 < 128 ? Color.White : Color.Black;
        }

        private static int GetFirstDataColumnIndex(DataGridView dgv)
        {
            for (int j = 0; j < dgv.Columns.Count; j++)
                if (dgv.Columns[j].Name != OriginalIndexColumnName) return j;
            return -1;
        }

        private static bool IsExpansionRow(DataGridView dgv, int rowIndex, int firstDataCol)
        {
            if (firstDataCol < 0 || rowIndex < 0 || rowIndex >= dgv.Rows.Count) return false;
            var val = dgv.Rows[rowIndex].Cells[firstDataCol].Value?.ToString()?.Trim() ?? "";
            return string.Equals(val, "Expansion", StringComparison.Ordinal);
        }

        private void ApplyRowStyles(DataGridView? grid = null)
        {
            var dgv = grid ?? dataGridView;
            if (dgv == null) return;

            int firstDataCol = GetFirstDataColumnIndex(dgv);
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (dgv.Rows[i].IsNewRow) continue;
                bool rowLocked = lockedRows.Contains(i);
                bool expansionRow = IsExpansionRow(dgv, i, firstDataCol);
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    if (dgv.Columns[j].Name == OriginalIndexColumnName) continue;
                    bool colLocked = lockedColumns.Contains(dgv.Columns[j].Index);
                    bool isFirstColumn = dgv.Columns[j].DisplayIndex == 0;
                    Color cellColor;
                    if (rowLocked || colLocked)
                        cellColor = _gridFrozenColor;
                    else if (expansionRow)
                        cellColor = _gridExpansionRowColor;
                    else if (i == 0 || isFirstColumn)
                        cellColor = _gridHeaderColor; // Row 0 = editable header row; first column = header column
                    else
                        cellColor = (i % 2 == 0 ? _gridEvenRowColor : _gridOddRowColor);
                    dgv.Rows[i].Cells[j].Style.BackColor = cellColor;
                    Color foreColor = (rowLocked || colLocked) ? TextColorCellFrozenActive : TextColorCellActive;
                    dgv.Rows[i].Cells[j].Style.ForeColor = foreColor;
                }
            }
        }

        private void SaveMenuItem_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentFilePath)) return;
            SaveFile(currentFilePath);
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private static string GetCurrentVersion()
        {
            var asm = Assembly.GetExecutingAssembly();
            var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            if (!string.IsNullOrEmpty(info?.InformationalVersion))
                return info.InformationalVersion.Split('+')[0].Trim();
            var ver = asm.GetName().Version;
            return ver != null ? $"{ver.Major}.{ver.Minor}.{ver.Build}" : "0.0.0";
        }

        private static bool IsNewerVersion(string current, string latest)
        {
            string Normalize(string v)
            {
                v = (v ?? "").Trim();
                if (v.StartsWith("v", StringComparison.OrdinalIgnoreCase)) v = v[1..];
                return v;
            }
            current = Normalize(current);
            latest = Normalize(latest);
            if (string.IsNullOrEmpty(latest)) return false;
            if (string.IsNullOrEmpty(current)) return true;
            int[] Parse(string s)
            {
                return s.Split('.', '-').Select(p => int.TryParse(p.Trim(), out int n) ? n : 0).ToArray();
            }
            var cur = Parse(current);
            var lat = Parse(latest);
            for (int i = 0; i < Math.Max(cur.Length, lat.Length); i++)
            {
                int c = i < cur.Length ? cur[i] : 0;
                int l = i < lat.Length ? lat[i] : 0;
                if (l > c) return true;
                if (l < c) return false;
            }
            return false;
        }

        private void OpenD2RDataguide()
        {
            try
            {
                string baseDir = AppContext.BaseDirectory ?? string.Empty;
                if (string.IsNullOrWhiteSpace(baseDir))
                {
                    MessageBox.Show(this, "Could not determine application directory.", "Open D2R Dataguide (2.9)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string indexPath = Path.Combine(baseDir, "index.html");
                if (!File.Exists(indexPath))
                {
                    MessageBox.Show(this, $"index.html was not found in:\n{baseDir}", "Open D2R Dataguide (2.9)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var psi = new ProcessStartInfo
                {
                    FileName = indexPath,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Failed to open D2R Dataguide:\n{ex.Message}", "Open D2R Dataguide (2.9)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void CheckForUpdatesAsync()
        {
            if (GitHubReleasesRepo.StartsWith("USERNAME/", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Update check is not configured. Set GitHubReleasesRepo in code to your repository (e.g. \"owner/repo\").", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string currentVersion = GetCurrentVersion();
            void SetChecking(bool checking)
            {
                this.Invoke(() =>
                {
                    if (_checkForUpdatesItem != null)
                    {
                        _checkForUpdatesItem.Enabled = !checking;
                        _checkForUpdatesItem.Text = checking ? "Checking for updates..." : "Check for &updates";
                    }
                });
            }
            SetChecking(true);
            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "D2_Horadrim-Updater/1.0");
                string url = $"https://api.github.com/repos/{GitHubReleasesRepo}/releases/latest";
                var response = await client.GetAsync(url).ConfigureAwait(true);
                if (!response.IsSuccessStatusCode)
                {
                    this.Invoke(() => MessageBox.Show($"Could not fetch releases: {response.StatusCode}. Check the repository path.", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                    return;
                }
                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                string tagName = root.TryGetProperty("tag_name", out var tagProp) ? tagProp.GetString() ?? "" : "";
                string htmlUrl = root.TryGetProperty("html_url", out var urlProp) ? urlProp.GetString() ?? "" : "";
                string latestVersion = tagName.TrimStart('v');
                if (string.IsNullOrWhiteSpace(latestVersion))
                {
                    this.Invoke(() => MessageBox.Show("Could not read latest release version.", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Warning));
                    return;
                }
                if (!IsNewerVersion(currentVersion, latestVersion))
                {
                    this.Invoke(() => MessageBox.Show($"You have the latest version ({currentVersion}).", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Information));
                    return;
                }
                string message = $"A new version ({latestVersion}) is available. You have {currentVersion}. Open the release page to download?";
                DialogResult result = DialogResult.None;
                this.Invoke(() => result = MessageBox.Show(message, "Update available", MessageBoxButtons.YesNo, MessageBoxIcon.Question));
                if (result == DialogResult.Yes && !string.IsNullOrEmpty(htmlUrl))
                {
                    try { Process.Start(new ProcessStartInfo(htmlUrl) { UseShellExecute = true }); }
                    catch (Exception ex) { MessageBox.Show($"Could not open link: {ex.Message}", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
            }
            catch (Exception ex)
            {
                this.Invoke(() => MessageBox.Show($"Update check failed: {ex.Message}", "Check for updates", MessageBoxButtons.OK, MessageBoxIcon.Warning));
            }
            finally
            {
                SetChecking(false);
            }
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopySelectedCellsToClipboard();
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PasteFromClipboard();
        }

        private void SaveFile(string filePath)
        {
            if (dataGridView == null || dataGridView.DataSource is not DataTable table) return;

            if (_createBackupOnSave && File.Exists(filePath))
            {
                int count = Math.Max(1, Math.Min(20, _backupFileCount));
                string basePath = filePath + ".bak";
                try
                {
                    if (count <= 1)
                    {
                        File.Copy(filePath, basePath, true);
                    }
                    else
                    {
                        if (File.Exists(basePath + "." + (count - 1)))
                            File.Delete(basePath + "." + (count - 1));
                        for (int i = count - 2; i >= 1; i--)
                        {
                            string oldName = basePath + "." + i;
                            string newName = basePath + "." + (i + 1);
                            if (File.Exists(newName)) File.Delete(newName);
                            if (File.Exists(oldName)) File.Move(oldName, newName);
                        }
                        string firstBackup = basePath + ".1";
                        if (File.Exists(firstBackup)) File.Delete(firstBackup);
                        if (File.Exists(basePath)) File.Move(basePath, firstBackup);
                        File.Copy(filePath, basePath, true);
                    }
                }
                catch { }
            }

            using (StreamWriter writer = new StreamWriter(filePath))
            {
                // Row 0 = header line, then data rows 1..Count-2 (last row may be new row)
                for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                {
                    var rowValues = new List<string>();
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        if (dataGridView.Columns[j].Name == OriginalIndexColumnName) continue;
                        rowValues.Add(dataGridView.Rows[i].Cells[j].Value?.ToString() ?? "");
                    }
                    writer.WriteLine(string.Join("\t", rowValues));
                }
            }
            _dirtyFilePaths.Remove(NormalizeFilePath(filePath));
            StoreSavedContent(filePath, GetContentStringFromTable(table));
            UpdateDirtyIndicators(filePath);
        }

        private void LoadDirectoryFiles()
        {
            if (string.IsNullOrEmpty(lastOpenedDirectory) || !Directory.Exists(lastOpenedDirectory)) return;

            string name = GetNextWorkspaceName();
            string key = name.Replace(" ", "_");
            if (treeView.Nodes.ContainsKey(key)) return;

            Color color = GetNextDefaultColor();
            var tag = new WorkspaceTag { FolderPath = lastOpenedDirectory, BackColor = color };
            TreeNode workspaceNode = new TreeNode(name) { Name = key, Tag = tag, BackColor = color };
            treeView.Nodes.Add(workspaceNode);
            foreach (var file in Directory.GetFiles(lastOpenedDirectory, "*.txt"))
                workspaceNode.Nodes.Add(new TreeNode(Path.GetFileName(file)) { Tag = file });
            ApplyTextColors();
        }

        private void TreeView_DrawNode(object? sender, DrawTreeNodeEventArgs e)
        {
            e.DrawDefault = false;
            Rectangle fullRow = new Rectangle(0, e.Bounds.Y, treeView.ClientSize.Width, e.Bounds.Height);

            bool isSelected = (e.State & TreeNodeStates.Selected) == TreeNodeStates.Selected;

            // Base colors (workspace nodes keep their custom background)
            Color baseBackColor = e.Node.Parent == null && e.Node.BackColor != Color.Empty
                ? e.Node.BackColor
                : treeView.BackColor;
            Color textColor = e.Node.ForeColor != Color.Empty ? e.Node.ForeColor : treeView.ForeColor;

            using (var brush = new SolidBrush(baseBackColor))
                e.Graphics.FillRectangle(brush, fullRow);

            const int expandSize = 9;
            int indent = e.Node.Level * (treeView.Indent);
            int expandX = 2 + indent;
            int expandY = e.Bounds.Y + (e.Bounds.Height - expandSize) / 2;
            var expandRect = new Rectangle(expandX, expandY, expandSize, expandSize);
            if (e.Node.Nodes.Count > 0)
            {
                ControlPaint.DrawBorder3D(e.Graphics, expandRect, Border3DStyle.RaisedInner);
                string glyph = e.Node.IsExpanded ? "−" : "+";
                using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                using (var brush = new SolidBrush(textColor))
                    e.Graphics.DrawString(glyph, treeView.Font, brush, expandRect, sf);
            }
            int textX = expandX + expandSize + 2;
            var textRect = new Rectangle(textX, e.Bounds.Y, treeView.ClientSize.Width - textX, e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, e.Node.Text, treeView.Font, textRect, textColor, baseBackColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPrefix);

            // Draw selection as an outline so workspace colors are never overridden.
            if (isSelected)
            {
                using (var pen = new Pen(SystemColors.Highlight, 1))
                {
                    var borderRect = new Rectangle(fullRow.X, fullRow.Y, fullRow.Width - 1, fullRow.Height - 1);
                    e.Graphics.DrawRectangle(pen, borderRect);
                }
            }
        }

        private void TreeView_MouseMove(object sender, MouseEventArgs e)
        {
            TreeNode hoveredNode = treeView.GetNodeAt(e.X, e.Y);
            if (hoveredNode != null && hoveredNode.Name.Contains("Workspace") && hoveredNode.Tag != null)
                toolTip1.SetToolTip(treeView, (hoveredNode.Tag as WorkspaceTag)?.FolderPath ?? hoveredNode.Tag.ToString() ?? "");
            else
                toolTip1.SetToolTip(treeView, string.Empty);
        }

        private void TreeView_BeforeLabelEdit(object? sender, NodeLabelEditEventArgs e)
        {
            // Only allow editing for workspace root nodes (no parent).
            // File nodes (children) must never be renamed from the workspace tree.
            if (e.Node?.Parent != null)
                e.CancelEdit = true;
        }

        private void TreeView_AfterLabelEdit(object? sender, NodeLabelEditEventArgs e)
        {
            if (e.Node?.Parent == null && e.Label != null)
                SaveWorkspacesConfig();
        }

        private void TreeView_ItemDrag(object? sender, ItemDragEventArgs e)
        {
            if (e.Item is TreeNode node && node.Parent == null)
                treeView.DoDragDrop(node, DragDropEffects.Move);
        }

        private void TreeView_DragOver(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(typeof(TreeNode)) != true)
            {
                e.Effect = DragDropEffects.None;
                HideDropIndicator();
                return;
            }
            e.Effect = DragDropEffects.Move;

            Point pt = treeView.PointToClient(new Point(e.X, e.Y));
            TreeNode? targetNode = treeView.GetNodeAt(pt);
            if (targetNode == null || targetNode.Parent != null)
            {
                HideDropIndicator();
                return;
            }
            if (e.Data?.GetData(typeof(TreeNode)) is TreeNode draggedNode && draggedNode == targetNode)
            {
                HideDropIndicator();
                return;
            }

            _dropTargetNode = targetNode;
            _dropInsertBefore = pt.Y < targetNode.Bounds.Y + targetNode.Bounds.Height / 2;
            ShowDropIndicator();
        }

        private void TreeView_DragLeave(object? sender, EventArgs e)
        {
            HideDropIndicator();
        }

        private void ShowDropIndicator()
        {
            if (_dropIndicatorPanel == null || _dropTargetNode == null || _dropIndicatorPanel.Parent == null) return;
            Rectangle r = _dropTargetNode.Bounds;
            int yInTree = _dropInsertBefore ? r.Y : r.Bottom;
            Point ptInPanel = _dropIndicatorPanel.Parent.PointToClient(treeView.PointToScreen(new Point(0, yInTree)));
            _dropIndicatorPanel.SetBounds(0, ptInPanel.Y - (_dropInsertBefore ? 2 : 0), treeView.Width, 3);
            _dropIndicatorPanel.Visible = true;
        }

        private void HideDropIndicator()
        {
            _dropTargetNode = null;
            if (_dropIndicatorPanel != null)
                _dropIndicatorPanel.Visible = false;
        }

        private void TreeView_DragDrop(object? sender, DragEventArgs e)
        {
            HideDropIndicator();

            if (e.Data?.GetData(typeof(TreeNode)) is not TreeNode draggedNode || draggedNode.Parent != null)
                return;

            Point pt = treeView.PointToClient(new Point(e.X, e.Y));
            TreeNode? targetNode = treeView.GetNodeAt(pt);
            if (targetNode == null || targetNode.Parent != null) return;
            if (draggedNode == targetNode) return;

            int draggedIndex = treeView.Nodes.IndexOf(draggedNode);
            int targetIndex = treeView.Nodes.IndexOf(targetNode);
            if (draggedIndex < 0 || targetIndex < 0) return;

            treeView.Nodes.Remove(draggedNode);
            int targetIndexAfter = (draggedIndex < targetIndex) ? targetIndex - 1 : targetIndex;
            int insertIndex = _dropInsertBefore ? targetIndexAfter : targetIndexAfter + 1;
            insertIndex = Math.Max(0, Math.Min(insertIndex, treeView.Nodes.Count));
            treeView.Nodes.Insert(insertIndex, draggedNode);
            treeView.SelectedNode = draggedNode;
            SaveWorkspacesConfig();
        }

        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string filePath && File.Exists(filePath))
            {
                currentFilePath = filePath;
                Color? wsColor = null;
                string? wsFolder = null;
                if (e.Node.Parent?.Tag is WorkspaceTag wsTag)
                {
                    wsColor = wsTag.BackColor;
                    wsFolder = wsTag.FolderPath;
                }
                LoadFile(filePath, wsColor, wsFolder);
            }
        }

        private void SaveLastDirectory()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    key.SetValue("LastDirectory", lastOpenedDirectory);
                }
            }
            catch { }
        }

        private void LoadLastDirectory()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key != null)
                        lastOpenedDirectory = key.GetValue("LastDirectory", "") as string;
                }
            }
            catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Application.AddMessageFilter(this);
            _lowLevelHookProc = LowLevelKeyboardHookCallback;
            _lowLevelHookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, _lowLevelHookProc, IntPtr.Zero, 0);
            try
            {
                var exeIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (exeIcon != null)
                    Icon = exeIcon;
            }
            catch { /* use default icon if extraction fails */ }
            Text = $"D2 Horadrim {GetCurrentVersion()}";
            string[] args = Environment.GetCommandLineArgs();
            ParseStartupArgs(args);
            LoadWindowSettings();
            LoadLastDirectory();
            LoadWorkspacesConfig();
            LoadExpandedWorkspaces();
            LoadGridColors();
            LoadTextColors();
            LoadCustomPaletteColors();
            ApplyTextColors();
            if (treeView.Nodes.Count == 0)
                LoadDirectoryFiles();
            if(_autoLoadPreviousSession && string.IsNullOrEmpty(_startupFileArgument))
                LoadOpenFiles();
            DisableColumnSorting();
        }

        private void ParseStartupArgs(string[] args)
        {
            _startupFileArgument = null;
            _startupWorkspaceRole = null;
            _startupWorkspaceFolder = null;
            if (args == null || args.Length <= 1) return;

            for (int i = 1; i < args.Length; i++)
            {
                string arg = args[i];
                if (string.IsNullOrWhiteSpace(arg)) continue;
                string lower = arg.ToLowerInvariant();

                bool isSource = lower.StartsWith("--source");
                bool isTarget = lower.StartsWith("--target");
                if (!isSource && !isTarget) continue;

                string role = isSource ? "Source" : "Target";

                // Find the --path flag (it may be combined, e.g. \"--target--path\")
                int pathFlagIndex = -1;
                if (lower.Contains("--path"))
                {
                    pathFlagIndex = i;
                }
                else
                {
                    for (int j = i + 1; j < args.Length; j++)
                    {
                        string candidate = args[j];
                        if (string.IsNullOrWhiteSpace(candidate)) continue;
                        string candLower = candidate.ToLowerInvariant();
                        if (candLower.Contains("--path"))
                        {
                            pathFlagIndex = j;
                            break;
                        }
                    }
                }

                if (pathFlagIndex >= 0)
                {
                    int folderIndex = pathFlagIndex + 1;
                    int fileIndex = pathFlagIndex + 2;
                    if (folderIndex < args.Length && fileIndex < args.Length)
                    {
                        string folderPath = args[folderIndex];
                        string fileNamePart = args[fileIndex];
                        if (!string.IsNullOrWhiteSpace(folderPath) && !string.IsNullOrWhiteSpace(fileNamePart))
                        {
                            _startupWorkspaceRole = role;
                            _startupWorkspaceFolder = folderPath;
                            _startupFileArgument = Path.Combine(folderPath, fileNamePart);
                            return;
                        }
                    }
                }
            }

            // Fallback to previous behavior if nothing matched
            if (args.Length > 1 && !string.IsNullOrWhiteSpace(args[1]))
                _startupFileArgument = args[1].Trim();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!_preserveD2CompareWorkspaces)
            {
                RemoveD2CompareWorkspaces();
            }

            SaveWindowSettings();
            SaveWorkspacesConfig();
            SaveOpenFiles();
            SaveExpandedWorkspaces();
            SaveGridColors();
            SaveTextColors();
            SaveCustomPaletteColors();
            SaveMyNotes();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (_lowLevelHookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_lowLevelHookHandle);
                _lowLevelHookHandle = IntPtr.Zero;
            }
            Application.RemoveMessageFilter(this);
            SaveWindowSettings();
        }

        private void LoadWindowSettings()
        {
            WindowState = FormWindowState.Maximized;
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key != null)
                    {
                        int x = (int)key.GetValue("WindowX", Location.X);
                        int y = (int)key.GetValue("WindowY", Location.Y);
                        int width = (int)key.GetValue("WindowWidth", Width);
                        int height = (int)key.GetValue("WindowHeight", Height);

                        StartPosition = FormStartPosition.Manual;
                        SetBounds(x, y, width, height);

                        int windowState = (int)key.GetValue("WindowState", 2);
                        WindowState = (FormWindowState)Math.Clamp(windowState, 0, 2);

                        int splitterDistance = (int)key.GetValue("SplitterDistance", 220);
                        int maxWorkspaceWidth = Math.Min(MaxWorkspacePanelWidth, Math.Max(220, (int)(Width * 0.28)));
                        splitContainer1.SplitterDistance = Math.Min(splitterDistance, maxWorkspaceWidth);

                        int editHistoryVisible = (int)key.GetValue("EditHistoryVisible", 1);
                        _editHistoryPaneVisible = editHistoryVisible != 0;
                        if (_editHistoryMenuItem != null)
                            _editHistoryMenuItem.Checked = _editHistoryPaneVisible;
                        int myNotesVisible = (int)key.GetValue("MyNotesVisible", 1);
                        _myNotesPaneVisible = myNotesVisible != 0;
                        if (_myNotesMenuItem != null)
                            _myNotesMenuItem.Checked = _myNotesPaneVisible;
                        int workspacePaneVisible = (int)key.GetValue("WorkspacePaneVisible", 1);
                        _workspacePaneVisible = workspacePaneVisible != 0;
                        if (_workspacePaneMenuItem != null)
                            _workspacePaneMenuItem.Checked = _workspacePaneVisible;
                        splitContainer1.Panel1Collapsed = !_workspacePaneVisible;
                        if (_workspaceSplitContainer != null && _bottomSplitContainer != null)
                        {
                            _bottomSplitContainer.Panel1Collapsed = !_editHistoryPaneVisible;
                            _bottomSplitContainer.Panel2Collapsed = !_myNotesPaneVisible;
                            _workspaceSplitContainer.Panel2Collapsed = !_editHistoryPaneVisible && !_myNotesPaneVisible;
                            if (_editHistoryPaneVisible || _myNotesPaneVisible)
                            {
                                int h = _workspaceSplitContainer.Height - _workspaceSplitContainer.SplitterWidth;
                                if (h > EditHistoryPaneHeight + 60)
                                {
                                    int wsSplit = (int)key.GetValue("WorkspaceSplitterDistance", h - EditHistoryPaneHeight);
                                    _workspaceSplitContainer.SplitterDistance = Math.Max(h - EditHistoryPaneHeight, Math.Min(wsSplit, h - 60));
                                }
                                int bottomSplit = (int)key.GetValue("BottomSplitterDistance", 0);
                                if (bottomSplit > 0 && _editHistoryPaneVisible && _myNotesPaneVisible)
                                    _bottomSplitContainer.SplitterDistance = Math.Max(80, Math.Min(bottomSplit, _bottomSplitContainer.Width - 80));
                            }
                        }
                        int darkMode = (int)key.GetValue("DarkMode", 1);
                        _darkMode = darkMode != 0;
                        if (_darkModeMenuItem != null)
                            _darkModeMenuItem.Checked = _darkMode;
                        int autoLoad = (int)key.GetValue("AutoLoadPreviousSession", 1);
                        _autoLoadPreviousSession = autoLoad != 0;
                        if (_autoLoadPreviousSessionMenuItem != null)
                            _autoLoadPreviousSessionMenuItem.Checked = _autoLoadPreviousSession;
                        int preserveD2Compare = SettingsStorage.GetSetting("PreserveD2CompareWorkspaces", 0);
                        _preserveD2CompareWorkspaces = preserveD2Compare != 0;
                        if (_preserveD2CompareWorkspacesMenuItem != null)
                            _preserveD2CompareWorkspacesMenuItem.Checked = _preserveD2CompareWorkspaces;
                        int headerTooltips = (int)key.GetValue("HeaderTooltipsFromDataGuide", 0);
                        _headerTooltipsFromDataGuide = headerTooltips != 0;
                        if (_headerTooltipsFromDataGuideMenuItem != null)
                            _headerTooltipsFromDataGuideMenuItem.Checked = _headerTooltipsFromDataGuide;
                        int alwaysLockHeaderCol = (int)key.GetValue("AlwaysLockHeaderColumn", 0);
                        _alwaysLockHeaderColumn = alwaysLockHeaderCol != 0;
                        if (_alwaysLockHeaderColumnMenuItem != null)
                            _alwaysLockHeaderColumnMenuItem.Checked = _alwaysLockHeaderColumn;
                        int alwaysLockHeaderRow = (int)key.GetValue("AlwaysLockHeaderRow", 0);
                        _alwaysLockHeaderRow = alwaysLockHeaderRow != 0;
                        if (_alwaysLockHeaderRowMenuItem != null)
                            _alwaysLockHeaderRowMenuItem.Checked = _alwaysLockHeaderRow;
                        int groupedColumnWidth = (int)key.GetValue("GroupedColumnWidth", 0);
                        if (_groupedColumnWidthMenuItem != null)
                            _groupedColumnWidthMenuItem.Checked = groupedColumnWidth != 0;
                        int groupedRowHeight = (int)key.GetValue("GroupedRowHeight", 0);
                        if (_groupedRowHeightMenuItem != null)
                            _groupedRowHeightMenuItem.Checked = groupedRowHeight != 0;
                        try
                        {
                            string? colJson = key.GetValue("SavedColumnWidthsByFile") as string;
                            if (!string.IsNullOrEmpty(colJson))
                            {
                                var deserialized = JsonSerializer.Deserialize<Dictionary<string, int[]>>(colJson);
                                _savedColumnWidthsByFile = deserialized != null ? new Dictionary<string, int[]>(deserialized, StringComparer.OrdinalIgnoreCase) : new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
                            }
                            string? rowJson = key.GetValue("SavedRowHeightsByFile") as string;
                            if (!string.IsNullOrEmpty(rowJson))
                            {
                                var deserialized = JsonSerializer.Deserialize<Dictionary<string, int[]>>(rowJson);
                                _savedRowHeightsByFile = deserialized != null ? new Dictionary<string, int[]>(deserialized, StringComparer.OrdinalIgnoreCase) : new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
                            }
                        }
                        catch { _savedColumnWidthsByFile = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase); _savedRowHeightsByFile = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase); }
                        int createBackup = (int)key.GetValue("CreateBackupOnSave", 0);
                        _createBackupOnSave = createBackup != 0;
                        int backupCount = (int)key.GetValue("BackupFileCount", 3);
                        _backupFileCount = Math.Max(1, Math.Min(20, backupCount));
                        string? fontName = key.GetValue("GridFontName") as string;
                        object? sizeObj = key.GetValue("GridFontSize");
                        float fontSize = sizeObj != null ? Convert.ToSingle(sizeObj) : 0f;
                        if (!string.IsNullOrEmpty(fontName) && fontSize > 0)
                        {
                            try
                            {
                                int styleVal = (int)key.GetValue("GridFontStyle", 0);
                                var style = (FontStyle)styleVal;
                                _gridFont = new Font(fontName, fontSize, style);
                            }
                            catch { _gridFont = null; }
                        }
                        else
                            _gridFont = null;
                    }
                    else
                    {
                        // No registry: use same defaults as when value is present
                        int maxWorkspaceWidth = Math.Min(MaxWorkspacePanelWidth, Math.Max(220, (int)(Width * 0.28)));
                        splitContainer1.SplitterDistance = Math.Min(220, maxWorkspaceWidth);
                        _editHistoryPaneVisible = true;
                        _myNotesPaneVisible = true;
                        if (_editHistoryMenuItem != null)
                            _editHistoryMenuItem.Checked = true;
                        if (_myNotesMenuItem != null)
                            _myNotesMenuItem.Checked = true;
                        _workspacePaneVisible = true;
                        if (_workspacePaneMenuItem != null)
                            _workspacePaneMenuItem.Checked = true;
                        splitContainer1.Panel1Collapsed = false;
                        if (_workspaceSplitContainer != null && _bottomSplitContainer != null)
                        {
                            _bottomSplitContainer.Panel1Collapsed = false;
                            _bottomSplitContainer.Panel2Collapsed = false;
                            _workspaceSplitContainer.Panel2Collapsed = false;
                            if (_workspaceSplitContainer.Height > EditHistoryPaneHeight + 60)
                            {
                                int h = _workspaceSplitContainer.Height - _workspaceSplitContainer.SplitterWidth;
                                _workspaceSplitContainer.SplitterDistance = Math.Min(h - EditHistoryPaneHeight, h - 60);
                                if (_bottomSplitContainer.Width > 160)
                                    _bottomSplitContainer.SplitterDistance = Math.Max(80, _bottomSplitContainer.Width / 2);
                            }
                        }
                    }
                }
            }
            catch { }
            ApplyTheme();
        }


        private void SaveWindowSettings()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    key.SetValue("WindowX", Location.X);
                    key.SetValue("WindowY", Location.Y);
                    key.SetValue("WindowWidth", Width);
                    key.SetValue("WindowHeight", Height);
                    key.SetValue("WindowState", (int)WindowState);
                    key.SetValue("SplitterDistance", Math.Min(splitContainer1.SplitterDistance, MaxWorkspacePanelWidth));
                    key.SetValue("EditHistoryVisible", _editHistoryPaneVisible ? 1 : 0);
                    key.SetValue("MyNotesVisible", _myNotesPaneVisible ? 1 : 0);
                    key.SetValue("WorkspacePaneVisible", _workspacePaneVisible ? 1 : 0);
                    if (_bottomSplitContainer != null)
                        key.SetValue("BottomSplitterDistance", _bottomSplitContainer.SplitterDistance);
                    key.SetValue("DarkMode", _darkMode ? 1 : 0);
                    key.SetValue("AutoLoadPreviousSession", _autoLoadPreviousSession ? 1 : 0);
                    key.SetValue("HeaderTooltipsFromDataGuide", _headerTooltipsFromDataGuide ? 1 : 0);
                    key.SetValue("AlwaysLockHeaderColumn", _alwaysLockHeaderColumn ? 1 : 0);
                    key.SetValue("AlwaysLockHeaderRow", _alwaysLockHeaderRow ? 1 : 0);
                    key.SetValue("GroupedColumnWidth", _groupedColumnWidthMenuItem?.Checked == true ? 1 : 0);
                    key.SetValue("GroupedRowHeight", _groupedRowHeightMenuItem?.Checked == true ? 1 : 0);
                    try
                    {
                        var colDict = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
                        var rowDict = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
                        foreach (TabPage tab in tabControl.TabPages)
                        {
                            string? filePath = (tab.Tag as TabTag)?.FilePath;
                            if (string.IsNullOrEmpty(filePath)) continue;
                            var dgv = GetGridForTab(tab);
                            if (dgv == null) continue;
                            string pathNorm = NormalizeFilePath(filePath);
                            if (string.IsNullOrEmpty(pathNorm)) continue;
                            var widths = new int[dgv.Columns.Count];
                            for (int i = 0; i < dgv.Columns.Count; i++)
                                widths[i] = dgv.Columns[i].Width;
                            colDict[pathNorm] = widths;
                            var heights = new int[dgv.Rows.Count];
                            for (int i = 0; i < dgv.Rows.Count; i++)
                                heights[i] = dgv.Rows[i].IsNewRow ? 0 : dgv.Rows[i].Height;
                            rowDict[pathNorm] = heights;
                        }
                        key.SetValue("SavedColumnWidthsByFile", JsonSerializer.Serialize(colDict));
                        key.SetValue("SavedRowHeightsByFile", JsonSerializer.Serialize(rowDict));
                    }
                    catch { }
                    key.SetValue("CreateBackupOnSave", _createBackupOnSave ? 1 : 0);
                    key.SetValue("BackupFileCount", Math.Max(1, Math.Min(20, _backupFileCount)));
                    if (_gridFont != null)
                    {
                        key.SetValue("GridFontName", _gridFont.Name);
                        key.SetValue("GridFontSize", _gridFont.SizeInPoints);
                        key.SetValue("GridFontStyle", (int)_gridFont.Style);
                    }
                    else
                    {
                        key.DeleteValue("GridFontName", false);
                        key.DeleteValue("GridFontSize", false);
                        key.DeleteValue("GridFontStyle", false);
                    }
                    if (_workspaceSplitContainer != null)
                        key.SetValue("WorkspaceSplitterDistance", _workspaceSplitContainer.SplitterDistance);
                }

                // Persist cross-platform settings via SettingsStorage as well.
                SettingsStorage.SetSetting("PreserveD2CompareWorkspaces", _preserveD2CompareWorkspaces ? 1 : 0);
            }
            catch { }
        }

        private void HeaderTooltipsFromDataGuideMenuItem_Click(object? sender, EventArgs e)
        {
            _headerTooltipsFromDataGuide = _headerTooltipsFromDataGuideMenuItem?.Checked ?? false;
        }

        private void AlwaysLockHeaderColumnMenuItem_Click(object? sender, EventArgs e)
        {
            _alwaysLockHeaderColumn = _alwaysLockHeaderColumnMenuItem?.Checked ?? false;
            SaveWindowSettings();
            if (dataGridView == null) return;
            int firstDataCol = GetFirstDataColumnIndex(dataGridView);
            if (firstDataCol < 0) return;
            if (_alwaysLockHeaderColumn)
                lockedColumns.Add(firstDataCol);
            else
                lockedColumns.Remove(firstDataCol);
            ApplyFrozenColumns();
            ApplyRowStyles();
            ApplyHeaderCellColors(dataGridView);
            RefreshColumnHeaders();
            dataGridView.Invalidate();
        }

        private void AlwaysLockHeaderRowMenuItem_Click(object? sender, EventArgs e)
        {
            _alwaysLockHeaderRow = _alwaysLockHeaderRowMenuItem?.Checked ?? false;
            SaveWindowSettings();
            if (dataGridView == null) return;
            if (_alwaysLockHeaderRow)
                lockedRows.Add(0);
            else
                lockedRows.Remove(0);
            ApplyFrozenRows();
            ApplyHeaderCellColors(dataGridView);
            RefreshColumnHeaders();
            dataGridView.Invalidate();
        }

        private void ConfigurationMenuItem_Click(object? sender, EventArgs e)
        {
            using (var form = new ConfigurationForm(_gridFont, _createBackupOnSave, _backupFileCount))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;
                _gridFont?.Dispose();
                _gridFont = form.SelectedFont != null ? (Font)form.SelectedFont.Clone() : null;
                _createBackupOnSave = form.CreateBackupOnSave;
                _backupFileCount = form.BackupFileCount;
                SaveWindowSettings();
                foreach (TabPage tab in tabControl.TabPages)
                {
                    var dgv = GetGridForTab(tab);
                    if (dgv != null && _gridFont != null)
                        dgv.Font = (Font)_gridFont.Clone();
                    else if (dgv != null)
                        dgv.Font = new Font(SystemFonts.DefaultFont.FontFamily, 9f);
                }
            }
        }

        // Shown via Help → Hotkeys.
        private static readonly string HotkeysLegend =
            "Row/Column must be selected:\r\n" +
            "Alt + F = Freeze/Unfreeze Row(s) or Column(s)\r\n" +
            "Alt + A = Add row or column\r\n" +
            "Alt + I = Insert row or column\r\n" +
            "Alt + R = Remove row or column\r\n" +
            "Alt + C = Clone row or column\r\n" +
            "Alt + H = Hide row or column\r\n" +
            "Alt + U = Unhide row or column\r\n\r\n" +
            "No selection needed:\r\n" +
            "Alt + E = Edit History pane toggle\r\n" +
            "Alt + N = Notes pane toggle\r\n" +
            "Alt + W = Workspace pane toggle\r\n" +
            "Alt + D = Dataguide Tooltips toggle";

        private static Dictionary<string, Dictionary<string, string>> BuildManualHeaderTooltips()
        {
            var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            var actinfo = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var armor = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var misc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var weapons = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var automagic = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var automap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var belts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var books = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var charstats = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cubemain = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var difficultylevels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var experience = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var gamble = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var gems = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var hireling = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var hirelingdesc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var inventory = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var itemratio = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var itemstatcost = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var itemtypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var levelgroups = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var levels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var lvlmaze = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var lvlprest = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var lvlsub = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var lvltypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var lvlwarp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var magicprefix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var magicsuffix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var missiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monequip = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monlvl = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monpreset = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monprop = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monseq = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monstats = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monstats2 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var montype = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monumod = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var monsounds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var npc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var objects = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var objgroup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var objpreset = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var overlay = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var pettype = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var qualityitems = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rareprefix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var raresuffix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var runes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var setitems = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var sets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var shrines = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var skilldesc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var skills = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var soundenviron = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var sounds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var states = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var superuniques = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var treasureclassex = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var uniqueappellation = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var uniqueitems = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var uniqueprefix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var uniquesuffix = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var wanderingmon = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            actinfo["act"] = "Open - Defines the ID for the Act.";
            actinfo["town"] = "Open - Uses an area level (Name field from Levels.txt) to define the Act's town area.";
            actinfo["start"] = "Open - Uses an area level (Name field from Levels.txt) to define where the player starts in the Act.";
            actinfo["maxnpcitemlevel"] = "Number - Controls the maximum item level for items sold by the NPC in the Act.";
            actinfo["classlevelrangestart"] = "Open - Uses an area level (Name field from Levels.txt) with its MonLvl values as a global Act minimum monster level. For example, this is used to determine chest levels in an Act.";
            actinfo["classlevelrangeend"] = "Open - Uses an area level (Name field from Levels.txt) with its MonLvl values as a global Act maximum monster level. For example, this is used to determine chest levels in an Act.";
            actinfo["wanderingnpcstart"] = "Number - Uses an index to determine which wandering monster class to use when populating areas (WanderingMon.txt for a list of possible monsters to spawn).";
            actinfo["wanderingnpcrange"] = "Number - This is a modifier that gets added to the wanderingnpcstart value to randomly select an index.";
            actinfo["commonactcof"] = "Open - Specifies which \".D2\" file to use as for the common Act COF file. This is used to establish the seed when initializing the Act.";
            actinfo["waypoint1"] = "Open - Uses an area level (Name field from Levels.txt) as the designated waypoint selection in the Waypoint UI.";
            actinfo["wanderingMonsterPopulateChance"] = "Number - The percent chance (from 0 to 100) to spawn a wandering monster (WanderingMon.txt for a list of possible monsters to spawn).";
            actinfo["wanderingMonsterRegionTotal"] = "Number - The maximum number of wandering monsters allowed at once.";
            actinfo["wanderingPopulateRandomChance"] = "Number - A secondary percent chance (from 0 to #) to determine whether to attempt populating with monsters. Only fails if random chance selects 0.";

            armor["name"] = "Open - This is a reference field to define the item; it does not need to be uniquely named.";
            armor["version"] = "Number - Defines which game version to create this item (0 = Classic mode | 100 = Expansion mode).";
            armor["compactsave"] = "Boolean - If equals 1, then only the item's base stats will be stored in the character save, but not any modifiers or additional stats. If equals 0, then all of the items stats will be saved.";
            armor["rarity"] = "Number - Determines the chance that the item will randomly spawn (1 / N). The higher the value then the rarer the item will be. This field depends on the spawnable field being enabled, the quest field being disabled, and the item level being less than or equal to the area level. This value is also affected by the relative Act number that the item is dropping in, where the higher the Act number, then the more common the item will drop.";
            armor["spawnable"] = "Boolean - If equals 1, then this item can be randomly spawned. If equals 0, then this item will never randomly spawn.";
            armor["speed"] = "Number - For Armors: Controls the Walk/Run Speed reduction when wearing the item. For Weapons: Controls the Attack Speed reduction when wearing the item.";
            armor["reqstr"] = "Number - Defines the amount of the Strength attribute needed to use the item.";
            armor["reqdex"] = "Number - Defines the amount of the Dexterity attribute needed to use the item.";
            armor["durability"] = "Number - Defines the base durability amount that the item will spawn with. An item must have durability to become ethereal.";
            armor["nodurability"] = "Boolean - If equals 1, then the item will not have durability. If equals 0, then the item will have durability.";
            armor["level"] = "Number - Controls the base item level. This is used for determining when the item is allowed to drop, such as making sure that the item level is not greater than the monster's level or the area level.";
            armor["ShowLevel"] = "Boolean - If equals 1, then display the item level next to the item name. If equals 0, then ignore this.";
            armor["levelreq"] = "Number - Controls the player level requirement for being able to use the item.";
            armor["cost"] = "Number - Defines the base gold cost of the item when being sold by an NPC. This can be affected by item modifiers and the rarity of the item.";
            armor["gamble cost"] = "Broken - Number - Defines the gambling gold cost of the item on the Gambling UI. Only functions for rings and amulets.";
            armor["code"] = "Open - Defines a unique 3* letter/number identifier code for the item. This code is case sensitive, so a69 and A69 are considered different. *Can technically use 4 characters instead of 3, but Items.json will not process it (no visuals).";
            armor["namestr"] = "Open - String Key that is used for the base item name. Reference string files found in local/lng/strings, such as item-names.json.";
            armor["alternategfx"] = "Open - Uses a unique 3 letter/number code similar to the defined code fields to determine what in-game graphics to display on the player character when the item is equipped.";
            armor["normcode"] = "Open - Links to the Code field to determine the normal version of the item.";
            armor["ubercode"] = "Open - Links to a Code field to determine the Exceptional version of the item.";
            armor["ultracode"] = "Open - Links to a Code field to determine the Elite version of the item.";
            armor["component"] = "Number - Determines the layer of player animation when the item is equipped. This uses a Token entry from Composit.txt.";
            armor["invwidth & invheight"] = "Number - Defines the width and height of grid cells that the item occupies in the player inventory";
            armor["hasinv"] = "Boolean - Determines if an item can be socketed with runes, jewels or Gems. (1 = Socketable | 0 = Not Socketable).";
            armor["Gemsockets"] = "Number - Controls the maximum number of sockets allowed on this item. This is limited by the item's size based on the invwidth and invheight fields. This also compares with the MaxSockets1 field(s) in the ItemTypes.txt file.";
            armor["gemapplytype"] = "Number - Determines which affect from a gem or rune will be applied when it is socketed into this item (Gems.txt).";
            armor["flippyfile"] = "Open - Controls which DC6 file to use for displaying the item in the game world when it is dropped on the ground (uses the file name as the input). *Legacy View Only.";
            armor["invfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory (uses the file name as the input). *Legacy View Only.";
            armor["uniqueinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Unique quality item (uses the file name as the input). *Legacy View Only (Any text required in field to use entry from uniques.JSON for D2R view).";
            armor["setinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Set quality item (uses the file name as the input). *Legacy View Only.";
            armor["useable"] = "Boolean - If equals 1, then the item can be used with the right-click mouse button command (this only works with specific belt items or quest items). If equals 0, then ignore this.";
            armor["stackable"] = "Boolean - If equals 1, then the item will use a quantity field and handle stacking functionality. This can depend on if the item type is throwable, is a type of ammunition, or is some other kind of miscellaneous item. If equals 0, then the item cannot be stacked.";
            armor["minstack"] = "Number - Controls the minimum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            armor["maxstack"] = "Number - Controls the maximum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            armor["spawnstack"] = "Number - Controls the stack count or quantity that the item can spawn with. This field depends on the stackable field being enabled.";
            armor["Transmogrify"] = "Broken - Boolean - If equals 1, then the item will use the transmogrify function. If equals 0, then ignore this. This field depends on the 'useable' field being enabled. Does not function at all";
            armor["TMogType"] = "Broken - Open - Links to a 'code' field to determine which item is chosen to transmogrify this item to. Does not function at all";
            armor["TMogMin"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            armor["TMogMax"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            armor["type"] = "Open - Points to an Code defined in the ItemTypes.txt file, which controls how the item functions.";
            armor["type2"] = "Open - Points to a secondary Item Type defined in the ItemTypes.txt file, which controls how the item functions. This is optional but can add more functionalities and possibilities with the item.";
            armor["dropsound"] = "Open - Points to a sound defined in the Sounds.txt file. Used when the item is dropped on the ground.";
            armor["dropsfxframe"] = "Number - Defines which frame in the flippyfile animation to play the dropsound when the item is dropped on the ground.";
            armor["usesound"] = "Open - Points to sound defined in the Sounds.txt file. Used when the item is moved in the inventory or used.";
            armor["unique"] = "Boolean - If equals 1, then the item can only spawn as a Unique quality type. If equals 0, then the item can spawn as other quality types.";
            armor["transparent"] = "Boolean - If equals 1, then the item will be drawn transparent on the player model (similar to ethereal models). If equals 0, then the item will appear solid on the player model.";
            armor["transtbl"] = "Number - Controls what type of transparency to use (if the above transparent field is enabled). Referenced by the Code value of the Transparency Table.";
            armor["lightradius"] = "Number - Controls the value of the light radius that this item can apply on the monster. This only affects monsters with this item equipped, not other types of units. This is ignored if the item's component on the monster is \"lit\", \"med\", or \"hvy\".";
            armor["belt"] = "Number - Controls which belt type to use for belt items only. This field determines what index entry in the Belts.txt file to use.";
            armor["quest"] = "Number - Controls what quest class is tied to the item which can enable certain item functionalities for a specific quest. Any value greater than 0 will also mean the item is flagged as a quest item, which can affect how it is displayed in tooltips, how it is traded with other players, its item rarity, and how it cannot be sold to an NPC. If equals 0, then the item will not be flagged as a quest item. Referenced by the Code value of the Quests Table.";
            armor["questdiffcheck"] = "Boolean - If equals 1 and the quest field is enabled, then the game will check the current difficulty setting and will tie that difficulty setting to the quest item. This means that the player can have more than 1 of the same quest item as long each they are obtained per difficulty mode (Normal / Nightmare / Hell). If equals 0 and the \"quest\" field is enabled, then the player can only have 1 count of the quest item in the inventory, regardless of difficulty.";
            armor["missiletype"] = "Open - Points to the *ID field from the Missiles.txt file, which determines what type of missile is used when using the throwing weapons.";
            armor["durwarning"] = "Number - Controls the threshold value for durability to display the low durability warning UI. This is only used if the item has durability.";
            armor["qntwarning"] = "Number - Controls the threshold value for quantity to display the low quantity warning UI. This is only used if the item has stacks.";
            armor["mindam"] = "Number - The minimum physical damage provided by the item.";
            armor["maxdam"] = "Number - The maximum physical damage provided by the item.";
            armor["StrBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Strength attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            armor["DexBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Dexterity attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            armor["gemoffset"] = "Number - Determines the starting index of possible gem/rune options based on the gemapplytype field and the Gems.txt entries. For example, if this value equals 9, then the game will start with index 9 (\"Chipped Emerald\") and ignore lower-indexed entries, meaning the item cannot roll these options.";
            armor["bitfield1"] = "Number - Controls different flags that can affect the item. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 5 (binary = 101) then that returns true for both the 4 (binary = 100) and 1 (binary = 1) bit field values.";
            armor["[NPC]Min"] = "Number - Minimum amount of this item type in Normal rarity that the NPC can sell at once";
            armor["[NPC]Max"] = "Number - Maximum amount of this item type in Normal rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            armor["[NPC]MagicMin"] = "Number - Minimum amount of this item type in Magical rarity that the NPC can sell at once";
            armor["[NPC]MagicMax"] = "Number - Maximum amount of this item type in Magical rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            armor["[NPC]MagicLvl"] = "Number - Maximum magic level allowed for this item type in Magical rarity; Where [NPC] is one of the following:";
            armor["Transform"] = "Open - Controls the color palette change of the item for the character model graphics. Referenced by the Code value of the Inventory Transform Table.";
            armor["InvTrans"] = "Number - Controls the color palette change of the item for the inventory graphics. Referenced by the Code value of the Inventory Transform Table.";
            armor["SkipName"] = "Boolean - If equals 1 and the item is Unique rarity, then skip adding the item's base name in its title. If equals 0, then ignore this.";
            armor["NightmareUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Nightmare difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            armor["HellUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Hell difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            armor["Nameable"] = "Boolean - If equals 1, then the item's name can be personalized by Anya for the Act 5 Betrayal of Harrogath quest reward. If equals 0, then the item cannot be used for the personalized name reward.";
            armor["PermStoreItem"] = "Boolean - If equals 1, then this item will always appear on the NPC's store. If equals 0, then the item will randomly appear on the NPC's store when appropriate.";
            armor["diablocloneweight"] = "Number - The amount of weight added to the diablo clone progress when this item is sold. When offline, selling this item will instead immediately spawn diablo clone.";
            armor["minac"] = "Number - The minimum amount of Defense that an armor item type can have.";
            armor["maxac"] = "Number - The maximum amount of Defense that an armor item type can have.";
            armor["block"] = "Number - Controls the block percent chance that the item provides (out of 100, but caps at 75).";
            armor["rArm"] = "Number - Controls the character's graphics and animations for the Right Arm component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";
            armor["lArm"] = "Number - Controls the character's graphics and animations for the Left Arm component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";
            armor["Torso"] = "Number - Controls the character's graphics and animations for the Torso component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";
            armor["Legs"] = "Number - Controls the character's graphics and animations for the Legs component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";
            armor["rSPad"] = "Number - Controls the character's graphics and animations for the Right Shoulder Pad component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";
            armor["lSPad"] = "Number - Controls the character's graphics and animations for the Left Shoulder Pad component when wearing the armor, where the value 0 = Light or \"lit\", 1 = Medium or \"med\", and 2 = Heavy or \"hvy\".";

            misc["name"] = "Open - This is a reference field to define the item; it does not need to be uniquely named.";
            misc["version"] = "Number - Defines which game version to create this item (0 = Classic mode | 100 = Expansion mode).";
            misc["compactsave"] = "Boolean - If equals 1, then only the item's base stats will be stored in the character save, but not any modifiers or additional stats. If equals 0, then all of the items stats will be saved.";
            misc["rarity"] = "Number - Determines the chance that the item will randomly spawn (1 / N). The higher the value then the rarer the item will be. This field depends on the spawnable field being enabled, the quest field being disabled, and the item level being less than or equal to the area level. This value is also affected by the relative Act number that the item is dropping in, where the higher the Act number, then the more common the item will drop.";
            misc["spawnable"] = "Boolean - If equals 1, then this item can be randomly spawned. If equals 0, then this item will never randomly spawn.";
            misc["speed"] = "Number - For Armors: Controls the Walk/Run Speed reduction when wearing the item. For Weapons: Controls the Attack Speed reduction when wearing the item.";
            misc["reqstr"] = "Number - Defines the amount of the Strength attribute needed to use the item.";
            misc["reqdex"] = "Number - Defines the amount of the Dexterity attribute needed to use the item.";
            misc["durability"] = "Number - Defines the base durability amount that the item will spawn with. An item must have durability to become ethereal.";
            misc["nodurability"] = "Boolean - If equals 1, then the item will not have durability. If equals 0, then the item will have durability.";
            misc["level"] = "Number - Controls the base item level. This is used for determining when the item is allowed to drop, such as making sure that the item level is not greater than the monster's level or the area level.";
            misc["ShowLevel"] = "Boolean - If equals 1, then display the item level next to the item name. If equals 0, then ignore this.";
            misc["levelreq"] = "Number - Controls the player level requirement for being able to use the item.";
            misc["cost"] = "Number - Defines the base gold cost of the item when being sold by an NPC. This can be affected by item modifiers and the rarity of the item.";
            misc["gamble cost"] = "Broken - Number - Defines the gambling gold cost of the item on the Gambling UI. Only functions for rings and amulets.";
            misc["code"] = "Open - Defines a unique 3* letter/number identifier code for the item. This code is case sensitive, so a69 and A69 are considered different. *Can technically use 4 characters instead of 3, but Items.json will not process it (no visuals).";
            misc["namestr"] = "Open - String Key that is used for the base item name. Reference string files found in local/lng/strings, such as item-names.json.";
            misc["alternategfx"] = "Open - Uses a unique 3 letter/number code similar to the defined code fields to determine what in-game graphics to display on the player character when the item is equipped.";
            misc["normcode"] = "Open - Links to the Code field to determine the normal version of the item.";
            misc["ubercode"] = "Open - Links to a Code field to determine the Exceptional version of the item.";
            misc["ultracode"] = "Open - Links to a Code field to determine the Elite version of the item.";
            misc["component"] = "Number - Determines the layer of player animation when the item is equipped. This uses a Token entry from Composit.txt.";
            misc["invwidth & invheight"] = "Number - Defines the width and height of grid cells that the item occupies in the player inventory";
            misc["hasinv"] = "Boolean - Determines if an item can be socketed with runes, jewels or Gems. (1 = Socketable | 0 = Not Socketable).";
            misc["Gemsockets"] = "Number - Controls the maximum number of sockets allowed on this item. This is limited by the item's size based on the invwidth and invheight fields. This also compares with the MaxSockets1 field(s) in the ItemTypes.txt file.";
            misc["gemapplytype"] = "Number - Determines which affect from a gem or rune will be applied when it is socketed into this item (Gems.txt).";
            misc["flippyfile"] = "Open - Controls which DC6 file to use for displaying the item in the game world when it is dropped on the ground (uses the file name as the input). *Legacy View Only.";
            misc["invfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory (uses the file name as the input). *Legacy View Only.";
            misc["uniqueinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Unique quality item (uses the file name as the input). *Legacy View Only (Any text required in field to use entry from uniques.JSON for D2R view).";
            misc["setinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Set quality item (uses the file name as the input). *Legacy View Only.";
            misc["useable"] = "Boolean - If equals 1, then the item can be used with the right-click mouse button command (this only works with specific belt items or quest items). If equals 0, then ignore this.";
            misc["stackable"] = "Boolean - If equals 1, then the item will use a quantity field and handle stacking functionality. This can depend on if the item type is throwable, is a type of ammunition, or is some other kind of miscellaneous item. If equals 0, then the item cannot be stacked.";
            misc["minstack"] = "Number - Controls the minimum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            misc["maxstack"] = "Number - Controls the maximum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            misc["spawnstack"] = "Number - Controls the stack count or quantity that the item can spawn with. This field depends on the stackable field being enabled.";
            misc["Transmogrify"] = "Broken - Boolean - If equals 1, then the item will use the transmogrify function. If equals 0, then ignore this. This field depends on the 'useable' field being enabled. Does not function at all";
            misc["TMogType"] = "Broken - Open - Links to a 'code' field to determine which item is chosen to transmogrify this item to. Does not function at all";
            misc["TMogMin"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            misc["TMogMax"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            misc["type"] = "Open - Points to an Code defined in the ItemTypes.txt file, which controls how the item functions.";
            misc["type2"] = "Open - Points to a secondary Item Type defined in the ItemTypes.txt file, which controls how the item functions. This is optional but can add more functionalities and possibilities with the item.";
            misc["dropsound"] = "Open - Points to a sound defined in the Sounds.txt file. Used when the item is dropped on the ground.";
            misc["dropsfxframe"] = "Number - Defines which frame in the flippyfile animation to play the dropsound when the item is dropped on the ground.";
            misc["usesound"] = "Open - Points to sound defined in the Sounds.txt file. Used when the item is moved in the inventory or used.";
            misc["unique"] = "Boolean - If equals 1, then the item can only spawn as a Unique quality type. If equals 0, then the item can spawn as other quality types.";
            misc["transparent"] = "Boolean - If equals 1, then the item will be drawn transparent on the player model (similar to ethereal models). If equals 0, then the item will appear solid on the player model.";
            misc["transtbl"] = "Number - Controls what type of transparency to use (if the above transparent field is enabled). Referenced by the Code value of the Transparency Table.";
            misc["lightradius"] = "Number - Controls the value of the light radius that this item can apply on the monster. This only affects monsters with this item equipped, not other types of units. This is ignored if the item's component on the monster is \"lit\", \"med\", or \"hvy\".";
            misc["belt"] = "Number - Controls which belt type to use for belt items only. This field determines what index entry in the Belts.txt file to use.";
            misc["quest"] = "Number - Controls what quest class is tied to the item which can enable certain item functionalities for a specific quest. Any value greater than 0 will also mean the item is flagged as a quest item, which can affect how it is displayed in tooltips, how it is traded with other players, its item rarity, and how it cannot be sold to an NPC. If equals 0, then the item will not be flagged as a quest item. Referenced by the Code value of the Quests Table.";
            misc["questdiffcheck"] = "Boolean - If equals 1 and the quest field is enabled, then the game will check the current difficulty setting and will tie that difficulty setting to the quest item. This means that the player can have more than 1 of the same quest item as long each they are obtained per difficulty mode (Normal / Nightmare / Hell). If equals 0 and the \"quest\" field is enabled, then the player can only have 1 count of the quest item in the inventory, regardless of difficulty.";
            misc["missiletype"] = "Open - Points to the *ID field from the Missiles.txt file, which determines what type of missile is used when using the throwing weapons.";
            misc["durwarning"] = "Number - Controls the threshold value for durability to display the low durability warning UI. This is only used if the item has durability.";
            misc["qntwarning"] = "Number - Controls the threshold value for quantity to display the low quantity warning UI. This is only used if the item has stacks.";
            misc["mindam"] = "Number - The minimum physical damage provided by the item.";
            misc["maxdam"] = "Number - The maximum physical damage provided by the item.";
            misc["StrBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Strength attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            misc["DexBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Dexterity attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            misc["gemoffset"] = "Number - Determines the starting index of possible gem/rune options based on the gemapplytype field and the Gems.txt entries. For example, if this value equals 9, then the game will start with index 9 (\"Chipped Emerald\") and ignore lower-indexed entries, meaning the item cannot roll these options.";
            misc["bitfield1"] = "Number - Controls different flags that can affect the item. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 5 (binary = 101) then that returns true for both the 4 (binary = 100) and 1 (binary = 1) bit field values.";
            misc["[NPC]Min"] = "Number - Minimum amount of this item type in Normal rarity that the NPC can sell at once";
            misc["[NPC]Max"] = "Number - Maximum amount of this item type in Normal rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            misc["[NPC]MagicMin"] = "Number - Minimum amount of this item type in Magical rarity that the NPC can sell at once";
            misc["[NPC]MagicMax"] = "Number - Maximum amount of this item type in Magical rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            misc["[NPC]MagicLvl"] = "Number - Maximum magic level allowed for this item type in Magical rarity; Where [NPC] is one of the following:";
            misc["Transform"] = "Open - Controls the color palette change of the item for the character model graphics. Referenced by the Code value of the Inventory Transform Table.";
            misc["InvTrans"] = "Number - Controls the color palette change of the item for the inventory graphics. Referenced by the Code value of the Inventory Transform Table.";
            misc["SkipName"] = "Boolean - If equals 1 and the item is Unique rarity, then skip adding the item's base name in its title. If equals 0, then ignore this.";
            misc["NightmareUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Nightmare difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            misc["HellUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Hell difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            misc["Nameable"] = "Boolean - If equals 1, then the item's name can be personalized by Anya for the Act 5 Betrayal of Harrogath quest reward. If equals 0, then the item cannot be used for the personalized name reward.";
            misc["PermStoreItem"] = "Boolean - If equals 1, then this item will always appear on the NPC's store. If equals 0, then the item will randomly appear on the NPC's store when appropriate.";
            misc["diablocloneweight"] = "Number - The amount of weight added to the diablo clone progress when this item is sold. When offline, selling this item will instead immediately spawn diablo clone.";
            misc["autobelt"] = "Boolean - If equals 1, then the item will automatically be placed is a free slot in the belt when picked up, if possible. If equals 0, then ignore this.";
            misc["bettergem"] = "Open - Links to another item's Code field. Also used by Function 18 from the Code field in Shrines.txt to specify which gem is the upgraded version.";
            misc["multibuy"] = "Boolean - If equals 1, then use the multi-buy transaction function when holding the shift key and buying this item from an NPC store. This multi-buy function will automatically purchase enough of the item to fill up to a full quantity stack or fill the available belt slots if the item is has the autobelt field enabled. If equals 0, then ignore this.";
            misc["spellicon"] = "Number - Determines the icon asset for displaying the item's spell. This uses an ID value based on the global skillicon file. If this value equals -1, then the item's spell will not display an icon. Used as a parameter for a PSpell function.";
            misc["pspell"] = "Number - Uses the Code value to select a spell from the Player Spell Table when the item is used. This depends on the item type.";
            misc["state"] = "Open - Links to a State field defined in the States.txt file. It signifies what state will be applied to the player when the item is used. Used as a parameter for a PSpell function.";
            misc["cstate1"] = "Open - Links to a State field defined in the States.txt file. It signifies what state will be removed from the player when the item is used. Used as a parameter for a PSpell function.";
            misc["len"] = "Number - Calculates the frame length of a state. Used as a parameter for a PSpell function.";
            misc["stat1"] = "Number - Controls the stat modifier when the item is used (Uses the Code field from Properties.txt). Used as a parameter for a PSpell function.";
            misc["calc1"] = "Calculation - Calculates the value of the above stat1 field. Used as a parameter for a PSpell function.";
            misc["spelldesc"] = "Number - Uses the code value from the below table to format and display string for the item's tooltip:.";
            misc["spelldescstr"] = "Open - Defines the primary string key used by the above spelldesc table.";
            misc["spelldescstr2"] = "Open - Defines the secondary string key used by the above spelldesc table.";
            misc["spelldesccalc"] = "Calculation - Value applied or used by the above spelldesc table.";
            misc["spelldesccolor"] = "Number - Uses the Code value from the below table, for the above spelldesc table.";

            weapons["name"] = "Open - This is a reference field to define the item; it does not need to be uniquely named.";
            weapons["version"] = "Number - Defines which game version to create this item (0 = Classic mode | 100 = Expansion mode).";
            weapons["compactsave"] = "Boolean - If equals 1, then only the item's base stats will be stored in the character save, but not any modifiers or additional stats. If equals 0, then all of the items stats will be saved.";
            weapons["rarity"] = "Number - Determines the chance that the item will randomly spawn (1 / N). The higher the value then the rarer the item will be. This field depends on the spawnable field being enabled, the quest field being disabled, and the item level being less than or equal to the area level. This value is also affected by the relative Act number that the item is dropping in, where the higher the Act number, then the more common the item will drop.";
            weapons["spawnable"] = "Boolean - If equals 1, then this item can be randomly spawned. If equals 0, then this item will never randomly spawn.";
            weapons["speed"] = "Number - For Armors: Controls the Walk/Run Speed reduction when wearing the item. For Weapons: Controls the Attack Speed reduction when wearing the item.";
            weapons["reqstr"] = "Number - Defines the amount of the Strength attribute needed to use the item.";
            weapons["reqdex"] = "Number - Defines the amount of the Dexterity attribute needed to use the item.";
            weapons["durability"] = "Number - Defines the base durability amount that the item will spawn with. An item must have durability to become ethereal.";
            weapons["nodurability"] = "Boolean - If equals 1, then the item will not have durability. If equals 0, then the item will have durability.";
            weapons["level"] = "Number - Controls the base item level. This is used for determining when the item is allowed to drop, such as making sure that the item level is not greater than the monster's level or the area level.";
            weapons["ShowLevel"] = "Boolean - If equals 1, then display the item level next to the item name. If equals 0, then ignore this.";
            weapons["levelreq"] = "Number - Controls the player level requirement for being able to use the item.";
            weapons["cost"] = "Number - Defines the base gold cost of the item when being sold by an NPC. This can be affected by item modifiers and the rarity of the item.";
            weapons["gamble cost"] = "Broken - Number - Defines the gambling gold cost of the item on the Gambling UI. Only functions for rings and amulets.";
            weapons["code"] = "Open - Defines a unique 3* letter/number identifier code for the item. This code is case sensitive, so a69 and A69 are considered different. *Can technically use 4 characters instead of 3, but Items.json will not process it (no visuals).";
            weapons["namestr"] = "Open - String Key that is used for the base item name. Reference string files found in local/lng/strings, such as item-names.json.";
            weapons["alternategfx"] = "Open - Uses a unique 3 letter/number code similar to the defined code fields to determine what in-game graphics to display on the player character when the item is equipped.";
            weapons["normcode"] = "Open - Links to the Code field to determine the normal version of the item.";
            weapons["ubercode"] = "Open - Links to a Code field to determine the Exceptional version of the item.";
            weapons["ultracode"] = "Open - Links to a Code field to determine the Elite version of the item.";
            weapons["component"] = "Number - Determines the layer of player animation when the item is equipped. This uses a Token entry from Composit.txt.";
            weapons["invwidth & invheight"] = "Number - Defines the width and height of grid cells that the item occupies in the player inventory";
            weapons["hasinv"] = "Boolean - Determines if an item can be socketed with runes, jewels or Gems. (1 = Socketable | 0 = Not Socketable).";
            weapons["Gemsockets"] = "Number - Controls the maximum number of sockets allowed on this item. This is limited by the item's size based on the invwidth and invheight fields. This also compares with the MaxSockets1 field(s) in the ItemTypes.txt file.";
            weapons["gemapplytype"] = "Number - Determines which affect from a gem or rune will be applied when it is socketed into this item (Gems.txt).";
            weapons["flippyfile"] = "Open - Controls which DC6 file to use for displaying the item in the game world when it is dropped on the ground (uses the file name as the input). *Legacy View Only.";
            weapons["invfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory (uses the file name as the input). *Legacy View Only.";
            weapons["uniqueinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Unique quality item (uses the file name as the input). *Legacy View Only (Any text required in field to use entry from uniques.JSON for D2R view).";
            weapons["setinvfile"] = "Open - Controls which DC6 file to use for displaying the item graphics in the inventory when it is a Set quality item (uses the file name as the input). *Legacy View Only.";
            weapons["useable"] = "Boolean - If equals 1, then the item can be used with the right-click mouse button command (this only works with specific belt items or quest items). If equals 0, then ignore this.";
            weapons["stackable"] = "Boolean - If equals 1, then the item will use a quantity field and handle stacking functionality. This can depend on if the item type is throwable, is a type of ammunition, or is some other kind of miscellaneous item. If equals 0, then the item cannot be stacked.";
            weapons["minstack"] = "Number - Controls the minimum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            weapons["maxstack"] = "Number - Controls the maximum stack count or quantity that is allowed on the item. This field depends on the stackable field being enabled.";
            weapons["spawnstack"] = "Number - Controls the stack count or quantity that the item can spawn with. This field depends on the stackable field being enabled.";
            weapons["Transmogrify"] = "Broken - Boolean - If equals 1, then the item will use the transmogrify function. If equals 0, then ignore this. This field depends on the 'useable' field being enabled. Does not function at all";
            weapons["TMogType"] = "Broken - Open - Links to a 'code' field to determine which item is chosen to transmogrify this item to. Does not function at all";
            weapons["TMogMin"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            weapons["TMogMax"] = "Broken - Number - Controls the minimum quantity that the transmogrify item will have. This depends on what item was chosen in the 'TMogType' field, and that the transmogrify item has quantity. Does not function at all";
            weapons["type"] = "Open - Points to an Code defined in the ItemTypes.txt file, which controls how the item functions.";
            weapons["type2"] = "Open - Points to a secondary Item Type defined in the ItemTypes.txt file, which controls how the item functions. This is optional but can add more functionalities and possibilities with the item.";
            weapons["dropsound"] = "Open - Points to a sound defined in the Sounds.txt file. Used when the item is dropped on the ground.";
            weapons["dropsfxframe"] = "Number - Defines which frame in the flippyfile animation to play the dropsound when the item is dropped on the ground.";
            weapons["usesound"] = "Open - Points to sound defined in the Sounds.txt file. Used when the item is moved in the inventory or used.";
            weapons["unique"] = "Boolean - If equals 1, then the item can only spawn as a Unique quality type. If equals 0, then the item can spawn as other quality types.";
            weapons["transparent"] = "Boolean - If equals 1, then the item will be drawn transparent on the player model (similar to ethereal models). If equals 0, then the item will appear solid on the player model.";
            weapons["transtbl"] = "Number - Controls what type of transparency to use (if the above transparent field is enabled). Referenced by the Code value of the Transparency Table.";
            weapons["lightradius"] = "Number - Controls the value of the light radius that this item can apply on the monster. This only affects monsters with this item equipped, not other types of units. This is ignored if the item's component on the monster is \"lit\", \"med\", or \"hvy\".";
            weapons["belt"] = "Number - Controls which belt type to use for belt items only. This field determines what index entry in the Belts.txt file to use.";
            weapons["quest"] = "Number - Controls what quest class is tied to the item which can enable certain item functionalities for a specific quest. Any value greater than 0 will also mean the item is flagged as a quest item, which can affect how it is displayed in tooltips, how it is traded with other players, its item rarity, and how it cannot be sold to an NPC. If equals 0, then the item will not be flagged as a quest item. Referenced by the Code value of the Quests Table.";
            weapons["questdiffcheck"] = "Boolean - If equals 1 and the quest field is enabled, then the game will check the current difficulty setting and will tie that difficulty setting to the quest item. This means that the player can have more than 1 of the same quest item as long each they are obtained per difficulty mode (Normal / Nightmare / Hell). If equals 0 and the \"quest\" field is enabled, then the player can only have 1 count of the quest item in the inventory, regardless of difficulty.";
            weapons["missiletype"] = "Open - Points to the *ID field from the Missiles.txt file, which determines what type of missile is used when using the throwing weapons.";
            weapons["durwarning"] = "Number - Controls the threshold value for durability to display the low durability warning UI. This is only used if the item has durability.";
            weapons["qntwarning"] = "Number - Controls the threshold value for quantity to display the low quantity warning UI. This is only used if the item has stacks.";
            weapons["mindam"] = "Number - The minimum physical damage provided by the item.";
            weapons["maxdam"] = "Number - The maximum physical damage provided by the item.";
            weapons["StrBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Strength attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            weapons["DexBonus"] = "Number - The percentage multiplier that gets multiplied the player's current Dexterity attribute value to modify the bonus damage percent from the equipped item. If this equals 1, then default the value to 100.";
            weapons["gemoffset"] = "Number - Determines the starting index of possible gem/rune options based on the gemapplytype field and the Gems.txt entries. For example, if this value equals 9, then the game will start with index 9 (\"Chipped Emerald\") and ignore lower-indexed entries, meaning the item cannot roll these options.";
            weapons["bitfield1"] = "Number - Controls different flags that can affect the item. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 5 (binary = 101) then that returns true for both the 4 (binary = 100) and 1 (binary = 1) bit field values.";
            weapons["[NPC]Min"] = "Number - Minimum amount of this item type in Normal rarity that the NPC can sell at once";
            weapons["[NPC]Max"] = "Number - Maximum amount of this item type in Normal rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            weapons["[NPC]MagicMin"] = "Number - Minimum amount of this item type in Magical rarity that the NPC can sell at once";
            weapons["[NPC]MagicMax"] = "Number - Maximum amount of this item type in Magical rarity that the NPC can sell at once. This must be equal to or greater than the minimum amount";
            weapons["[NPC]MagicLvl"] = "Number - Maximum magic level allowed for this item type in Magical rarity; Where [NPC] is one of the following:";
            weapons["Transform"] = "Open - Controls the color palette change of the item for the character model graphics. Referenced by the Code value of the Inventory Transform Table.";
            weapons["InvTrans"] = "Number - Controls the color palette change of the item for the inventory graphics. Referenced by the Code value of the Inventory Transform Table.";
            weapons["SkipName"] = "Boolean - If equals 1 and the item is Unique rarity, then skip adding the item's base name in its title. If equals 0, then ignore this.";
            weapons["NightmareUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Nightmare difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            weapons["HellUpgrade"] = "Open - Links to another item's Code field. Used to determine which item will replace this item when being generated in the NPC's store while the game is playing in Hell difficulty. If this field's code equals \"xxx\", then this item will not change in this difficulty.";
            weapons["Nameable"] = "Boolean - If equals 1, then the item's name can be personalized by Anya for the Act 5 Betrayal of Harrogath quest reward. If equals 0, then the item cannot be used for the personalized name reward.";
            weapons["PermStoreItem"] = "Boolean - If equals 1, then this item will always appear on the NPC's store. If equals 0, then the item will randomly appear on the NPC's store when appropriate.";
            weapons["diablocloneweight"] = "Number - The amount of weight added to the diablo clone progress when this item is sold. When offline, selling this item will instead immediately spawn diablo clone.";
            weapons["1or2handed"] = "Boolean - If equals 1, then the item will be treated as a one-handed and two-handed weapon by the Barbarian class. If equals 0, then the Barbarian can only use this weapon as either one-handed or two-handed, but not both.";
            weapons["2handed"] = "Boolean - If equals 1, then the item will be treated as two-handed weapon. If equals 0, then the item will be treated as one-handed weapon.";
            weapons["2handedwclass"] = "Open - Defines the two-handed weapon class, which controls what character animations are used when the weapon is equipped. Referenced by the Code value of the Weapon Class Table.";
            weapons["2handmindam"] = "Number - The minimum physical damage provided by the weapon if the item is two-handed. This relies on the 2handed field being enabled.";
            weapons["2handmaxdam"] = "Number - The maximum physical damage provided by the weapon if the item is two-handed. This relies on the 2handed field being enabled.";
            weapons["minmisdam"] = "Number - The maximum physical damage provided by the item if it is a throwing weapon.";
            weapons["maxmisdam"] = "Number - The maximum physical damage provided by the item if it is a throwing weapon.";
            weapons["rangeadder"] = "Number - Adds extra range in grid spaces for melee attacks while the melee weapon is equipped. The baseline melee range is 1, and this field adds to that range.";
            weapons["wclass"] = "Open - Defines the one-handed weapon class, which controls what character animations are used when the weapon is equipped. Referenced by the Code value of the Weapon Class Table.";

            automagic["Name"] = "Open - Defines the item affix name; it does not need to be uniquely named.";
            automagic["version"] = "Number - Defines which game version to use this item affix (0 = Classic mode | 100 = Expansion mode).";
            automagic["spawnable"] = "Boolean - If equals 1, then this item affix is used as part of the game's randomizer for assigning item modifiers when an item spawns. If equals 0, then this item affix is never used.";
            automagic["rare"] = "Boolean - If equals 1, then this item affix can be used when randomly assigning item modifiers when a rare item spawns. If equals 0, then this item affix is not used for rare items.";
            automagic["level"] = "Number - The minimum item level required for this item affix to spawn on the item. If the item level is below this value, then the item affix will not spawn on the item.";
            automagic["maxlevel"] = "Number - The maximum item level required for this item affix to spawn on the item. If the item level is above this value, then the item affix will not spawn on the item.";
            automagic["levelreq"] = "Number - The minimum character level required to equip an item that has this item affix.";
            automagic["classspecific"] = "Open - Controls if this item affix should only be used for class specific items. Specified by the Code column in PlayerClass.txt.";
            automagic["class"] = "Open - Controls which character class is required for the class specific level requirement specified in the below classlevelreq field. Specified by the Code column in PlayerClass.txt.";
            automagic["classlevelreq"] = "Number - The minimum character level required for a specific class in order to equip an item that has this item affix. This relies on the class specified in the above class field. If equals null, then the class will default to using the levelreq field.";
            automagic["frequency"] = "Number - Controls the probability that the affix appears on the item (a higher value means that the item affix will appear on the item more often). This value gets summed together with other frequency values from all possible item affixes that can spawn on the item, and then is used as a denominator value for the randomizer. Whichever item affix is randomly selected will be the one to appear on the item. The formula is calculated as the following: [Item Affix Selected] = [\"frequency\"] / [Total Frequency]. If the item has a magic lvl, then the magic level value is multiplied with this value. If equals 0, then this item affix will never appear on an item.";
            automagic["group"] = "Number - Assigns an item affix to a specific group number. Items cannot spawn with more than 1 item affix with the same group number. This is used to guarantee that certain item affixes do not overlap on the same item. If this field is null, then the group number will default to group 0.";
            automagic["mod1code"] = "Open - Controls the item properties for the item affix using the Code value from Properties.txt.";
            automagic["mod1param"] = "Number - The \"parameter\" value associated with the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            automagic["mod1min"] = "Number - The \"min\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            automagic["mod1max"] = "Number - The \"max\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            automagic["transformcolor"] = "Number - Controls the color change of the item after spawning with this item affix. If empty, then the item affix will not change the item's color. Referenced from the Code column in colors.txt.";
            automagic["itype1"] = "Open - Controls what Item Types are allowed to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            automagic["etype1"] = "Open - Controls what Item Types are forbidden to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            automagic["multiply"] = "Number - Multiplicative modifier for the item's buy and sell costs, based on the item affix (Calculated in 1024ths for buy cost and 4096ths for sell cost).";
            automagic["add"] = "Number - Flat integer modification to the item's buy and sell costs, based on the item affix.";

            automap["LevelName"] = "Open - Uses a string format system to define the Act number and name of the level type. Level types are static defined values that cannot be added. The number at the start of the string defines the Act number, and the word that follows this number defines the level type. This data should stay grouped by level.";
            automap["TileName"] = "Open - Uses defined string codes to control the tile orientations on the Automap.";
            automap["Style"] = "Number - Defines a group numeric ID for the range of cells, meaning that the game will try to use cells that matc the same style value, after determining the Level Type and Tile Type. If this value is equal to 255, then the style is ignored in the \"Cel#\" field selection.";
            automap["StartSequence"] = "Number - The start index value for valid Cel1 field(s) to choose for displaying on the Automap. If this value is equal to 255, then both the \"StartSequence\" and \"EndSequence\" are ignored in the Cel1 field(s) field selection. If this value is equal to -1, then this field is ignored in the Cel1 field(s) field selection.";
            automap["EndSequence"] = "Number - The end index value for a valid Cel1 field(s) to choose for displaying on the Automap. If this value is equal to -1, then this field is ignored in the Cel1 field(s) selection.";
            automap["Cel1"] = "Number - Determines the unique image frame to use from the MaxiMap.dc6 file that will be used to display on the Automap for that position of the level tile. There are multiple of these fields because they can be randomly chosen to give image variety in the Automap display. If the value equals -1, then this cell is not valid and will be ignored. If no cell is chosen overall, then nothing will be drawn in this area on the Automap.";

            belts["name"] = "Open - This is a reference field to define the belt type.";
            belts["numboxes"] = "Number - This integer field defines the number of item slots in the belt. This is used when inserting items into the belt and also for handling the removal of items when the belt item is unequipped.";
            belts["box1left"] = "Number - Specifies the belt slot left side coordinates. This is use for Server verification purposes and does not affect the local box UI in the client.";
            belts["box1right"] = "Number - Specifies the belt slot right side coordinates. This is use for Server verification purposes and does not affect the local box UI in the client.";
            belts["box1top"] = "Number - Specifies the belt slot left top coordinates. This is use for Server verification purposes and does not affect the local box UI in the client.";
            belts["box1bottom"] = "Number - Specifies the belt slot bottom side coordinates. This is use for Server verification purposes and does not affect the local box UI in the client.";
            belts["defaultItemTypeCol1"] = "Open - Specifies the default item type used for the populate belt and auto-use functionality on controller.";
            belts["defaultItemCodeCol1"] = "Open - Specifies the default item code used for the populate belt and auto-use functionality on controller. Leaving this blank uses no code and instead relies entirely on the item type.";

            books["Name"] = "Open - This is a reference field to define the book.";
            books["ScrollSpellCode"] = "Open - Uses an item's code to define as the scroll item for the book.";
            books["BookSpellCode"] = "Open - Uses an item's code to define as the book item.";
            books["pSpell"] = "Number - Defines the item spell function to use when using the book. Referenced by the Code value of the Player Spell Table.";
            books["SpellIcon"] = "Open - Controls which DC6 file to display for the mouse cursor when using the scroll or book (Uses numeric indices to pick the DC6 file. Example: When using Identify, use icon 1 or buysell.DC6).";
            books["ScrollSkill"] = "Open - Defines which Skill to use for the scroll item (Uses the skill field from Skills.txt).";
            books["BookSkill"] = "Open - Defines which Skill to use for the book item (Uses the skill field from Skills.txt).";
            books["BaseCost"] = "Number - The starting gold cost to buy the book from an NPC.";
            books["CostPerCharge"] = "Number - The additional gold cost added with the book's \"BaseCost\" value, based on how many charges the book has.";

            charstats["class"] = "Open - The name of the character class (this cannot be changed).";
            charstats["str"] = "Number - Starting amount of the Strength attribute.";
            charstats["dex"] = "Number - Starting amount of the Dexterity attribute.";
            charstats["int"] = "Number - Starting amount of the Energy attribute.";
            charstats["vit"] = "Number - Starting amount of the Vitality attribute.";
            charstats["stamina"] = "Number - Starting amount of Stamina.";
            charstats["hpadd"] = "Number - Bonus starting Life value (This value gets added with the vit field value to determine the overall starting amount of Life).";
            charstats["ManaRegen"] = "Number - Number of seconds to regain max Mana. (If this equals 0 then it will default to 300 seconds).";
            charstats["ToHitFactor"] = "Number - Starting amount of Attack Rating.";
            charstats["WalkVelocity"] = "Number - Base Walk movement speed.";
            charstats["RunVelocity"] = "Number - Base Run movement speed.";
            charstats["RunDrain"] = "Number - Rate at which Stamina is lost while running.";
            charstats["LifePerLevel"] = "Number - Amount of Life added for each level gained (Calculated in fourths and is divided by 256).";
            charstats["StaminaPerLevel"] = "Number - Amount of Stamina added for each level gained (Calculated in fourths and is divided by 256).";
            charstats["LifePerVitality"] = "Number - Amount of Life added for each point in Vitality (Calculated in fourths and is divided by 256).";
            charstats["StaminaPerVitality"] = "Number - Amount of Stamina added for each point in Vitality (Calculated in fourths and is divided by 256).";
            charstats["ManaPerMagic"] = "Number - Amount of Mana added for each point in Energy (Calculated in fourths and is divided by 256).";
            charstats["StatPerLevel"] = "Number - Amount of Attribute stat points earned for each level gained.";
            charstats["SkillsPerLevel"] = "Number - Amount of Skill points earned for each level gained.";
            charstats["LightRadius"] = "Number - Baseline radius size of the character's Light Radius.";
            charstats["BlockFactor"] = "Number - Baseline percent chance for Blocking.";
            charstats["MinimumCastingDelay"] = "Number - Global delay on all Skills after using a Skill with a Casting Delay (Calculated in Frames, where 25 Frames = 1 Second).";
            charstats["StartSkill"] = "Open - Controls what skill will be added by default to the character's starting weapon and will be slotted in the Right Skill selection (Uses the skill field from Skills.txt).";
            charstats["Skill 1"] = "Open - Skill that the character starts with and will always have available (Uses the skill field from Skills.txt).";
            charstats["StrAllSkills"] = "Open - String key for displaying the item modifier bonus to all skills for the class (Ex: \"+1 to Barbarian Skill Levels\").";
            charstats["StrSkillTab1"] = "Open - String key for displaying the item modifier bonus to all skills for the class's first to third skill tab (Ex: \"+1 to Warcries\").";
            charstats["StrClassOnly"] = "Open - String key for displaying on item modifier exclusive to the class or for class specific items (Ex: \"Barbarian only\").";
            charstats["HealthPotionPercent"] = "Number - This scales the amount of Life that a Healing potion will restore based on the class.";
            charstats["ManaPotionPercent"] = "Number - This scales the amount of Mana that a Mana potion will restore based on the class.";
            charstats["baseWClass"] = "Open - Base weapon class that the character will use by default when no weapon is equipped. Referenced by the Code value of the Weapon Class Table.";
            charstats["item1"] = "Open - Item that the character starts with (Uses the code field from Weapons.txt, Armor.txt or Misc.txt).";
            charstats["item1loc"] = "Open - Location where the related item will be placed in the character's inventory. Referenced from the Code column in BodyLocs.txt.";
            charstats["item1count"] = "Number - The amount of the related item that the character starts with.";
            charstats["item1quality"] = "Open - Controls the quality level of the related item using the below table:.";

            cubemain["description"] = "Open - This is a reference field to define the cube recipe.";
            cubemain["enabled"] = "Boolean - If equals 1, then the recipe can be used in-game. If equals 0, then the recipe cannot be used in-game.";
            cubemain["ladder"] = "Boolean - If equals 1, then the recipe can only be used on Ladder realm games. If equals 0, then the recipe can be used in all game types.";
            cubemain["min diff"] = "Number - The minimum game difficulty to use the recipe (0 = All Game Difficulties | 1 = Nightmare and Hell Difficulty only | 2 = Hell Difficulty only).";
            cubemain["version"] = "Number - Defines which game version to use this recipe (0 = Classic mode | 100 = Expansion mode).";
            cubemain["op"] = "Number - Uses a function as an additional input requirement for the recipe. (See param and value).";
            cubemain["param"] = "Number - Integer value used as a possible parameter for the op function. Normally uses the *ID field from ItemStatCost.txt for the param value.";
            cubemain["value"] = "Number - Integer value used as a possible parameter for the op function. Normally used to compare against the *ID linked in param.";
            cubemain["class"] = "Open - Defines the recipe to be only usable by a defined class. Referenced from the Code column in PlayerClass.txt.";
            cubemain["numinputs"] = "Number - Controls the number of items that need to be inside the cube for the recipe.";
            cubemain["input 1"] = "Open - Controls what items are required for the recipe. Uses either the item's unique code or a combination of pre-defined codes from the table below, separated by commas if needed. (hiq,noe).";
            cubemain["output"] = "Open - Controls the first to third output items. Uses either the item's unique code or a combination of pre-defined codes from the table below, separated by commas if needed. (hiq,noe).";
            cubemain["lvl"] = "Number - Forces the output item level to be a specific level. If this field is used, then ignore the below plvl and ilvl fields.";
            cubemain["plvl"] = "Number - This is a numeric ratio that gets multiplied with the current player's level, to add to the output item's level requirement.";
            cubemain["ilvl"] = "Number - This is a numeric ratio that gets multiplied with input 1's item's level, to add to the output item's level requirement.";
            cubemain["mod 1"] = "Open - Controls the output item properties (Uses the Code field from Properties.txt).";
            cubemain["mod 1 chance"] = "Number - The percent chance that the property will be assigned. If this equals 0, then the Property will always be assigned.";
            cubemain["mod 1 param"] = "Number - The \"parameter\" value associated with the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            cubemain["mod 1 min"] = "Number - The \"min\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            cubemain["mod 1 max"] = "Number - The \"max\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";

            difficultylevels["Name"] = "Open - This is a reference field to define the difficulty mode.";
            difficultylevels["ResistPenalty"] = "Number - Defines the baseline starting point for a player character's resistances for Expansion mode.";
            difficultylevels["ResistPenaltyNonExpansion"] = "Number - Defines the baseline starting point for a player character's resistances for Non-Expansion mode.";
            difficultylevels["DeathExpPenalty"] = "Number - Modifies the percentage of current level Experience lost when a player character dies.";
            difficultylevels["MonsterSkillBonus"] = "Number - Adds additional skill levels to skills used by monsters (Only applies to monsters with skills assigned via the Skill1 column(s) in MonStats.txt).";
            difficultylevels["MonsterFreezeDivisor"] = "Number - Divisor that affects all Freeze Length values on monsters. The attempted Freeze Length value is divided by this divisor to determine the actual Freeze Length.";
            difficultylevels["MonsterColdDivisor"] = "Number - Divisor that affects all Cold Length values on monsters. The attempted Cold Length value is divided by this divisor to determine the actual Cold Length.";
            difficultylevels["AiCurseDivisor"] = "Number - Divisor that affects all durations of Curses on monsters. The attempted Curse duration is divided by this divisor to determine the actual Curse duration.";
            difficultylevels["LifeStealDivisor"] = "Number - Divisor that affects the amount of Life Steal that player characters gain. The attempted Life Steal value is divided by this divisor to determine the actual Life Steal.";
            difficultylevels["ManaStealDivisor"] = "Number - Divisor that affects the amount of Mana Steal that player characters gain. The attempted Mana Steal value is divided by this divisor to determine the actual Mana Steal.";
            difficultylevels["UniqueDamageBonus"] = "Number - Percentage modifier for a Unique monster's overall Damage and Attack Rating. This is applied after calculating the monster's other modifications.";
            difficultylevels["ChampionDamageBonus"] = "Number - Percentage modifier for a Champion monster's overall Damage and Attack Rating. This is applied after calculating the monster's other modifications.";
            difficultylevels["PlayerDamagePercentVSPlayer"] = "Number - Percentage modifier for the total damage a player deals to another player.";
            difficultylevels["PlayerDamagePercentVSMercenary"] = "Number - Percentage modifier for the total damage a player deals to another player's mercenary.";
            difficultylevels["PlayerDamagePercentVSPrimeEvil"] = "Number - Percentage modifier for the total damage a player deals to a Prime Evil boss.";
            difficultylevels["PlayerHitReactBufferVSPlayer"] = "Number - The frame length for the amount of time a player cannot be placed into another hit react from a player (25 frames = 1 second).";
            difficultylevels["PlayerHitReactBufferVSMonster"] = "Number - The frame length for the amount of time a player cannot be placed into another hit react from a monster (25 frames = 1 second).";
            difficultylevels["MercenaryDamagePercentVSPlayer"] = "Number - Percentage modifier for the total damage a player's mercenary deals to another player.";
            difficultylevels["MercenaryDamagePercentVSMercenary"] = "Number - Percentage modifier for the total damage a player's mercenary deals to another player's mercenary.";
            difficultylevels["MercenaryDamagePercentVSBoss"] = "Number - Percentage modifier for the total damage a player's mercenary deals to a boss monster.";
            difficultylevels["MercenaryMaxStunLength"] = "Number - The frame length for the maximum stun length allowed on a player's mercenary (25 Frames = 1 second).";
            difficultylevels["PrimeEvilDamagePercentVSPlayer"] = "Number - Percentage modifier applied to the total damage a Prime Evil boss deals to a player.";
            difficultylevels["PrimeEvilDamagePercentVSMercenary"] = "Number - Percentage modifier for the total damage a Prime Evil boss deals to a player's mercenary.";
            difficultylevels["PrimeEvilDamagePercentVSPet"] = "Number - Percentage modifier for the total damage a Prime Evil boss deals to a player's pet.";
            difficultylevels["PetDamagePercentVSPlayer"] = "Number - Percentage modifier for the total damage a player's pet deals to another player.";
            difficultylevels["MonsterCEDamagePercent"] = "Number - Percentage modifier that affects how much damage is dealt to a player by a Monster's version of Corpse Explosion. For example, when certain monsters die and explode on death.";
            difficultylevels["MonsterFireEnchantExplosionDamagePercent"] = "Number - Percentage modifier that affects how much damage is dealt to a player by a Monster's Fire Enchant explosion. The Fire Enchant death explosion uses the same Corpse Explosion functionality and this value is applied after the \"MonsterCEDamagePercent\" value.";
            difficultylevels["StaticFieldMin"] = "Number - Percentage modifier for capping the amount of current Life damage dealt to monsters by the Sorceress Static Field skill. This field only affects games in Expansion mode.";
            difficultylevels["GambleRare"] = "Number - The odds to obtain a Rare item from gambling. The game rolls a random number between 0 to 100000. If that rolled number is less than this value, then the gambled item will be a Rare item.";
            difficultylevels["GambleSet"] = "Number - The odds to obtain a Set item from gambling. The game rolls a random number between 0 to 100000. If that rolled number is less than this value, then the gambled item will be a Set item.";
            difficultylevels["GambleUnique"] = "Number - The odds to obtain a Unique item from gambling. The game rolls a random number between 0 to 100000. If that rolled number is less than this value, then the gambled item will be a Unique item.";
            difficultylevels["GambleUber"] = "Number - The odds to make the gambled item be an Exceptional Quality item. The game rolls a random number between 0 to 10000. This rolled number is then compared to the following formula:.";
            difficultylevels["GambleUltra"] = "Number - The odds to make the gambled item be an Elite Quality item. The game rolls a random number between 0 to 10000. This rolled number is then compared to the following formula:.";

            experience["Level"] = "Open - This is a reference field to define the level.";
            experience["Amazon"] = "Number - Controls the Experience required for each level with the Amazon class.";
            experience["Sorceress"] = "Number - Controls the Experience required for each level with the Sorceress class.";
            experience["Necromancer"] = "Number - Controls the Experience required for each level with the Necromancer class.";
            experience["Paladin"] = "Number - Controls the Experience required for each level with the Paladin class.";
            experience["Barbarian"] = "Number - Controls the Experience required for each level with the Barbarian class.";
            experience["Druid"] = "Number - Controls the Experience required for each level with the Druid class.";
            experience["Assassin"] = "Number - Controls the Experience required for each level with the Assassin class.";
            experience["ExpRatio"] = "Number - This multiplier affects the percentage of Experienced earned, based on the level (Calculated in 1024ths by default, but this can be changed by updating the \"10\" value in the \"Maxlvl\" row).";

            gamble["name"] = "Open - This is a reference field to describe the Item.";
            gamble["code"] = "Open - This is a pointer to code field from Weapons.txt/Armor.txt/Misc.txt.";

            gems["name"] = "Open - This is a reference field to define the gem/rune name.";
            gems["letter"] = "Open - Defines the string that is concatenated together in the item tooltip when a rune is socketed into an item.";
            gems["transform"] = "Number - Controls the color change of the item after being socketed by the gem/rune. Referenced from the Index value in colors.txt.";
            gems["code"] = "Open - Defines the unique item code used to create the gem/rune.";
            gems["weaponMod1Code"] = "Open - Controls the item properties that the gem/rune provides when socketed into an item with a \"gemapplytype\" value that equals 0 (Uses the Code field from Properties.txt).";
            gems["weaponMod1Param"] = "Number - The stat's \"parameter\" value associated with the listed property (weaponMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["weaponMod1Min"] = "Number - The stat's \"min\" value associated with the listed property (weaponMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["weaponMod1Max"] = "Number - The stat's \"max\" value to assign to the listed property (weaponMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["helmMod1Code"] = "Open - Controls the item properties that the gem/rune provides when socketed into an item with a \"gemapplytype\" value that equals 1 (Uses the Code field from Properties.txt).";
            gems["helmMod1Param"] = "Number - The stat's \"parameter\" value associated with the listed property (helmMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["helmMod1Min"] = "Number - The stat's \"min\" value associated with the listed property (helmMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["helmMod1Max"] = "Number - The stat's \"max\" value to assign to the listed property (helmMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["shieldMod1Code"] = "Open - Controls the item properties that the gem/rune provides when socketed into an item with a \"gemapplytype\" value that equals 2 (Uses the Code field from Properties.txt).";
            gems["shieldMod1Param"] = "Number - The stat's \"parameter\" value associated with the listed property (shieldMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["shieldMod1Min"] = "Number - The stat's \"min\" value associated with the listed property (shieldMod1Code). Usage depends on the (Function ID field from Properties.txt).";
            gems["shieldMod1Max"] = "Number - The stat's \"max\" value to assign to the listed property (shieldMod1Code). Usage depends on the (Function ID field from Properties.txt).";

            hireling["Hireling"] = "Open - This is a reference field to define the Hireling name.";
            hireling["Version"] = "Number - Defines which game version to use this hireling (0 = Classic mode | 100 = Expansion mode).";
            hireling["ID"] = "Number - The unique identification number to define each hireling type.";
            hireling["Class"] = "Number - This refers to the hcIDx field in MonStats.txt, which defines the base type of unit to use for the hireling.";
            hireling["Act"] = "Number - The Act that the hireling belongs to (values 1 to 5 equal Act 1 to Act 5, respectively).";
            hireling["Difficulty"] = "Number - The difficulty mode associated with the hireling (1 = Normal | 2 = Nightmare | 3 = Hell).";
            hireling["Level"] = "Number - The starting level of the unit.";
            hireling["Seller"] = "Number - This refers to the hcIDx field in MonStats.txt, which defines the unit NPC that sells this hireling.";
            hireling["NameFirst & NameLast"] = "Open - These fields define a string key which the game uses as a sequential range of string IDs from 'NameFirst' to 'NameLast' to randomly generate as hireling names. (Max name length is 48 characters)";
            hireling["Gold"] = "Number - The initial cost of the hireling. This is used in the following calculation to generate the full hire price: Cost = [\"Gold\"] * (100 + 15 * [Difference of Current Level and \"Level\"]) / 100.";
            hireling["Exp/Lvl"] = "Number - This modifier is used in the following calculation to determine the amount of Experience need for the hireling's next level: [Current Level] + [Current Level] * [Current Level + 1] * ['Exp / Lvl']";
            hireling["HP"] = "Number - The starting amount of Life at base Level.";
            hireling["HP/Lvl"] = "Number - The amount of Life gained per Level";
            hireling["Defense"] = "Number - The starting amount of Defense at base Level.";
            hireling["Def/Lvl"] = "Number - The amount of Defense gained per Level";
            hireling["Str"] = "Number - The starting amount of Strength at base Level.";
            hireling["Str/Lvl"] = "Number - The amount of Strength gained per Level (Calculated in 8ths)";
            hireling["Dex"] = "Number - The starting amount of Dexterity at base Level.";
            hireling["Dex/Lvl"] = "Number - The amount of Dexterity gained per Level (Calculated in 8ths)";
            hireling["AR"] = "Number - The starting amount of Attack Rating at base Level.";
            hireling["AR/Lvl"] = "Number - The amount of Attack Rating gained per Level";
            hireling["Dmg-Min"] = "Number - The starting amount of minimum Physical Damage for attacks";
            hireling["Dmg-Max"] = "Number - The starting amount of maximum Physical Damage for attacks";
            hireling["Dmg/Lvl"] = "Number - The amount of Physical Damage gained per level, to be added to 'Dmg - Min' and 'Dmg - Max' (Calculated in 8ths)";
            hireling["ResistFire"] = "Number - The starting amount of Fire Resistance at base Level.";
            hireling["ResistFire/Lvl"] = "Number - The amount of Fire Resistance gained per Level (Calculated in 4ths)";
            hireling["ResistCold"] = "Number - The starting amount of Fire Resistance at base Level.";
            hireling["ResistCold/Lvl"] = "Number - The amount of Fire Resistance gained per Level (Calculated in 4ths)";
            hireling["ResistLightning"] = "Number - The starting amount of Fire Resistance at base Level.";
            hireling["ResistLightning/Lvl"] = "Number - The amount of Fire Resistance gained per Level (Calculated in 4ths)";
            hireling["ResistPoison"] = "Number - The starting amount of Fire Resistance at base Level.";
            hireling["ResistPoison/Lvl"] = "Number - The amount of Fire Resistance gained per Level (Calculated in 4ths)";
            hireling["HireDesc"] = "Open - This accepts a string key, which is used to display as the special description of the hireling in the hire UI window (local/lng/strings/mercenaries.JSON).";
            hireling["DefaultChance"] = "Number - This is the chance for the hireling to attack with his/her weapon instead of using a Skill. All Chance values are summed together as a denominator value for a random roll to determine which skill to use.";
            hireling["Skill1"] = "Open - Points to a skill from the skill field in the Skills.txt file. This gives the hireling the Skill to use (requires \"Mode#\", \"Chance#\", \"ChancePerLvl#\").";
            hireling["Mode1"] = "Open - Uses a monster mode to determine the hireling's behavior when using the related Skill. Referenced from the Index value in MonMode.txt.";
            hireling["Chance1"] = "Number - This is the base chance for the hireling to use the related Skill. All Chance values are summed together as a denominator value for a random roll to determine which skill to use.";
            hireling["ChancePerLvl1"] = "Number - This is the chance for the hireling to use the related Skill, affected by the difference in the hireling's current Level and the hireling's Level field. All Chance values are summed together as a denominator value for a random roll to determine which skill to use. Each skill Chance is calculated with the following formula: [\"Chance#\"] + [\"ChancePerLvl#\"] * [Difference of Current Level and \"Level\"] / 4.";
            hireling["Level1"] = "Number - The starting Level for the related Skill.";
            hireling["LvlPerLvl1"] = "Number - A modifier to increase the related Skill level for every Level gained. This is used in the following calculated to determine the current skill level: [Current Skill Level] = FLOOR([\"Level\"] + (([\"LvlPerLvl\"] * [Difference of Current Level and \"Level\"]) / 32)).";
            hireling["HiringMaxLevelDifference"] = "Number - This is used to generate a range with this value plus and minus with the player's current Level. In the hiring UI window, hirelings start with a random Level that is between this range.";
            hireling["resurrectcostmultiplier"] = "Number - A modifier used to calculate the hireling's current resurrect cost. Used in the following formula: [Resurrect Cost] = [Current Level] * [Current Level] / [\"resurrectcostdivisor\"] * [\"resurrectcostmultiplier\"].";
            hireling["resurrectcostdivisor"] = "Number - A modifier used to calculate the hireling's current resurrect cost. Used in the following formula: [Resurrect Cost] = [Current Level] * [Current Level] / [\"resurrectcostdivisor\"] * [\"resurrectcostmultiplier\"].";
            hireling["resurrectcostmax"] = "Number - This is the maximum Gold cost to resurrect this hireling.";
            hireling["equivalentcharclass"] = "Open - Determines what class this hireling is treated like under the hood when calculating skill level bonuses and gear restrictions.";

            hirelingdesc["id"] = "Number - The id of the hireling monster class as defined in monstats.txt.";
            hirelingdesc["alternateVoice"] = "Boolean - If equals 1, then the hireling will use the alternate (feminine) voice type for voice lines. If equals 0, then it will use the masculine voice type.";

            inventory["class"] = "Open - This is a reference field to define the type of inventory screen.";
            inventory["invLeft"] = "Number - Starting X coordinate pixel position of the inventory panel.";
            inventory["invRight"] = "Number - Ending X coordinate pixel position of the inventory panel (Includes the \"invLeft\" value with the inventory width size).";
            inventory["invTop"] = "Number - Starting Y coordinate pixel position of the inventory panel.";
            inventory["invBottom"] = "Number - Ending Y coordinate pixel position of the inventory panel (Includes the \"invTop\" value with the inventory height size).";
            inventory["gridX"] = "Number - Column number size of the inventory grid, measured in the number of grid boxes to use.";
            inventory["gridY"] = "Number - Column row size of the inventory grid, measured in the number of grid boxes to use.";
            inventory["gridLeft"] = "Number - Starting X coordinate location of the inventory's left grid side.";
            inventory["gridRight"] = "Number - Ending X coordinate location of the inventory's right grid side (Includes the \"gridLeft\" value with the grid width size).";
            inventory["gridTop"] = "Number - Starting Y coordinate location of the inventory's top grid side.";
            inventory["gridBottom"] = "Number - Ending Y coordinate location of the inventory's bottom grid side (Includes the \"gridTop\" value with the grid height size).";
            inventory["gridBoxWidth"] = "Number - Width size of an inventory's box cell.";
            inventory["gridBoxHeight"] = "Number - Height size of an inventory's box cell.";
            inventory["rArmLeft"] = "Number - Starting X coordinate location of the Right Weapon Slot.";
            inventory["rArmRight"] = "Number - Ending X coordinate location of the Right Weapon Slot (Includes the \"rArmLeft\" value with the \"rArmWidth\" value).";
            inventory["rArmTop"] = "Number - Starting Y coordinate location of the Right Weapon Slot.";
            inventory["rArmBottom"] = "Number - Ending Y coordinate location of the Right Weapon Slot (Includes the \"rArmTop\" value with the \"rArmHeight\" value).";
            inventory["rArmWidth"] = "Number - The pixel width of the Right Weapon Slot.";
            inventory["rArmHeight"] = "Number - The pixel Height of the Right Weapon Slot.";
            inventory["torsoLeft"] = "Number - Starting X coordinate location of the Body Armor Slot.";
            inventory["torsoRight"] = "Number - Ending X coordinate location of the Body Armor Slot (Includes the \"torsoLeft\" value with the \"torsoWidth\" value).";
            inventory["torsoTop"] = "Number - Starting Y coordinate location of the Body Armor Slot.";
            inventory["torsoBottom"] = "Number - Ending Y coordinate location of the Body Armor Slot (Includes the \"torsoTop\" value with the \"torsoHeight\" value).";
            inventory["torsoWidth"] = "Number - The pixel width of the Body Armor Slot.";
            inventory["torsoHeight"] = "Number - The pixel Height of the Body Armor Slot.";
            inventory["lArmLeft"] = "Number - Starting X coordinate location of the Left Weapon Slot.";
            inventory["lArmRight"] = "Number - Ending X coordinate location of the Left Weapon Slot (Includes the \"lArmLeft\" value with the \"lArmWidth\" value).";
            inventory["lArmTop"] = "Number - Starting Y coordinate location of the Left Weapon Slot.";
            inventory["lArmBottom"] = "Number - Ending Y coordinate location of the Left Weapon Slot (Includes the \"lArmTop\" value with the \"lArmHeight\" value).";
            inventory["lArmWidth"] = "Number - The pixel width of the Left Weapon Slot.";
            inventory["lArmHeight"] = "Number - The pixel Height of the Left Weapon Slot.";
            inventory["headLeft"] = "Number - Starting X coordinate location of the Helm Slot.";
            inventory["headRight"] = "Number - Ending X coordinate location of the Helm Slot (Includes the \"headLeft\" value with the \"headWidth\" value).";
            inventory["headTop"] = "Number - Starting Y coordinate location of the Helm Slot.";
            inventory["headBottom"] = "Number - Ending Y coordinate location of the Helm Slot (Includes the \"headTop\" value with the \"headHeight\" value).";
            inventory["headWidth"] = "Number - The pixel width of the Helm Slot.";
            inventory["headHeight"] = "Number - The pixel Height of the Helm Slot.";
            inventory["neckLeft"] = "Number - Starting X coordinate location of the Amulet Slot.";
            inventory["neckRight"] = "Number - Ending X coordinate location of the Amulet Slot (Includes the \"neckLeft\" value with the \"neckWidth\" value).";
            inventory["neckTop"] = "Number - Starting Y coordinate location of the Amulet Slot.";
            inventory["neckBottom"] = "Number - Ending Y coordinate location of the Amulet Slot (Includes the \"neckTop\" value with the \"neckHeight\" value).";
            inventory["neckWidth"] = "Number - The pixel width of the Amulet Slot.";
            inventory["neckHeight"] = "Number - The pixel Height of the Amulet Slot.";
            inventory["rHandLeft"] = "Number - Starting X coordinate location of the Right Ring Slot.";
            inventory["rHandLeft"] = "Number - Ending X coordinate location of the Right Ring Slot (Includes the \"rHandLeft\" value with the \"rHandWidth\" value).";
            inventory["rHandTop"] = "Number - Starting Y coordinate location of the Right Ring Slot.";
            inventory["rHandBottom"] = "Number - Ending Y coordinate location of the Right Ring Slot (Includes the \"rHandTop\" value with the \"rHandHeight\" value).";
            inventory["rHandWidth"] = "Number - The pixel width of the Right Ring Slot.";
            inventory["rHandHeight"] = "Number - The pixel Height of the Right Ring Slot.";
            inventory["lHandLeft"] = "Number - Starting X coordinate location of the Left Ring Slot.";
            inventory["lHandRight"] = "Number - Ending X coordinate location of the Left Ring Slot (Includes the \"lHandLeft\" value with the \"lHandWidth\" value).";
            inventory["lHandTop"] = "Number - Starting Y coordinate location of the Left Ring Slot.";
            inventory["lHandBottom"] = "Number - Ending Y coordinate location of the Left Ring Slot (Includes the \"lHandTop\" value with the \"lHandHeight\" value).";
            inventory["lHandWidth"] = "Number - The pixel width of the Left Ring Slot.";
            inventory["lHandHeight"] = "Number - The pixel Height of the Left Ring Slot.";
            inventory["beltLeft"] = "Number - Starting X coordinate location of the Belt Slot.";
            inventory["beltRight"] = "Number - Ending X coordinate location of the Belt Slot (Includes the \"beltLeft\" value with the \"beltWidth\" value).";
            inventory["beltTop"] = "Number - Starting Y coordinate location of the Belt Slot.";
            inventory["beltBottom"] = "Number - Ending Y coordinate location of the Belt Slot (Includes the \"beltTop\" value with the \"beltHeight\" value).";
            inventory["beltWidth"] = "Number - The pixel width of the Belt Slot.";
            inventory["beltHeight"] = "Number - The pixel Height of the Belt Slot.";
            inventory["feetLeft"] = "Number - Starting X coordinate location of the Boots Slot.";
            inventory["feetRight"] = "Number - Ending X coordinate location of the Boots Slot (Includes the \"feetLeft\" value with the \"feetWidth\" value).";
            inventory["feetTop"] = "Number - Starting Y coordinate location of the Boots Slot.";
            inventory["feetBottom"] = "Number - Ending Y coordinate location of the Boots Slot (Includes the \"feetTop\" value with the \"feetHeight\" value).";
            inventory["feetWidth"] = "Number - The pixel width of the Boots Slot.";
            inventory["feetHeight"] = "Number - The pixel Height of the Boots Slot.";
            inventory["glovesLeft"] = "Number - Starting X coordinate location of the Gloves Slot.";
            inventory["glovesRight"] = "Number - Ending X coordinate location of the Gloves Slot (Includes the \"glovesLeft\" value with the \"glovesWidth\" value).";
            inventory["glovesTop"] = "Number - Starting Y coordinate location of the Gloves Slot.";
            inventory["glovesBottom"] = "Number - Ending Y coordinate location of the Gloves Slot (Includes the \"glovesTop\" value with the \"glovesHeight\" value).";
            inventory["glovesWidth"] = "Number - The pixel width of the Gloves Slot.";
            inventory["glovesHeight"] = "Number - The pixel Height of the Gloves Slot.";

            itemratio["Function"] = "Open - This is a reference field to define the item ratio name.";
            itemratio["Version"] = "Number - Defines which game version to use this item ratio (0 = Classic mode | 100 = Expansion mode).";
            itemratio["Uber"] = "Boolean - If equals 1, then the item ratio will apply to items with Exceptional or Elite Quality. If equals 0, then the item ratio will apply to Normal Quality items (This is determined by the normcode, ubercode and ultracode fields in Armor/Weapons.txt).";
            itemratio["Class Specific"] = "Boolean - If equals 1, then the item ratio will apply to class-based items (This will compare to the Item Type's Class field from ItemTypes.txt to determine if the item is class specific).";
            itemratio["Unique"] = "Number - Base value for calculating the Unique Quality chance. Higher value means rarer chance. (Calculated first).";
            itemratio["UniqueDivisor"] = "Number - Modifier for changing the Unique Quality chance, based on the difference between the Monster Level and the Item's base level.";
            itemratio["UniqueMin"] = "Number - The minimum value of the probability denominator for Unique Quality. This is compared to the calculated Unique Quality value after Magic Find calculations and is chosen if it is greater than that value. (Calculated in 128ths).";
            itemratio["Set"] = "Number - Base value for calculating the Set Quality chance. Higher value means rarer chance. (Calculated after Unique).";
            itemratio["SetDivisor"] = "Number - Modifier for changing the Set Quality chance, based on the difference between the Monster Level and the Item's base level.";
            itemratio["SetMin"] = "Number - The minimum value of the probability denominator for Set Quality. This is compared to the calculated Set Quality value after Magic Find calculations and is chosen if it is greater than that value. (Calculated in 128ths).";
            itemratio["Rare"] = "Number - Base value for calculating the Rare Quality chance. Higher value means rarer chance. (Calculated after Set).";
            itemratio["RareDivisor"] = "Number - Modifier for changing the Rare Quality chance, based on the difference between the Monster Level and the Item's base level.";
            itemratio["RareMin"] = "Number - The minimum value of the probability denominator for Rare Quality. This is compared to the calculated Rare Quality value after Magic Find calculations and is chosen if it is greater than that value. (Calculated in 128ths).";
            itemratio["Magic"] = "Number - Base value for calculating the Magic Quality chance. Higher value means rarer chance. (Calculated after Rare).";
            itemratio["MagicDivisor"] = "Number - Modifier for changing the Magic Quality chance, based on the difference between the Monster Level and the Item's base level.";
            itemratio["MagicMin"] = "Number - The minimum value of the probability denominator for Magic Quality. This is compared to the calculated Magic Quality value after Magic Find calculations and is chosen if it is greater than that value. (Calculated in 128ths).";
            itemratio["HiQuality"] = "Number - Base value for calculating the High Quality (Superior) chance. Higher value means rarer chance. (Calculated after Magic).";
            itemratio["HiQualityDivisor"] = "Number - Modifier for changing the High Quality (Superior) chance, based on the difference between the Monster Level and the Item's base level.";
            itemratio["Normal"] = "Number - Base value for calculating the Normal Quality chance. Higher value means rarer chance. (Calculated after Normal, and if this does not succeed in rolling, then the item is defaulted to Low Quality).";
            itemratio["NormalDivisor"] = "Number - Modifier for changing the Normal Quality chance, based on the difference between the Monster Level and the Item's base level.";

            itemstatcost["Stat"] = "Open - Defines the unique pointer for this stat, which is used in other files.";
            itemstatcost["Send Other"] = "Boolean - If equals 1, then only add the stat to a new monster if the that has no state and has an item mask. If equals 0, then ignore this.";
            itemstatcost["Signed"] = "Boolean - If equals 1, then the stat will be treated as a signed integer, meaning that it can be either a positive or negative value. If equals 0, then stat will be treated as an unsigned integer, meaning that it can only be a positive value. This only affects stats with state bits.";
            itemstatcost["Send Bits"] = "Number - Controls how many bits of data for the stat to send to the game client, essentially controlling the max value possible for the stat. Signed values should have less than 32 bits, otherwise they will be treated as unsigned values.";
            itemstatcost["Send Param Bits"] = "Number - Controls how many bits of data for the stat's parameter value to send to the client for a unit. This value is always treated as a signed integer.";
            itemstatcost["UpdateAnimRate"] = "Boolean - If equals 1, then the stat will notify that game to handle and adjust the speed of the unit when the stat changes. If equals 0, then ignore this. This is only checked for stats with States or for specific skill server functions including 30, 61, 71.";
            itemstatcost["Saved"] = "Boolean - If equals 1, then this state will be inserted in the change list to be stored in the Character Save file. If equals 0, then ignore this.";
            itemstatcost["CSvSigned"] = "Boolean - If equals 1, then the stat will be saved as a signed integer in the Character Save file. If equals 0, then the stat will be saved as an unsigned integer in the Character Save file. This is only used if the \"Saved\" field is enabled.";
            itemstatcost["CSvBits"] = "Number - Controls how many bits of data for the stat to send to save in the Character Save file. Signed values should have less than 32 bits, otherwise they will be treated as unsigned values. This is only used if the \"Saved\" field is enabled.";
            itemstatcost["CSvParam"] = "Number - Controls how many bits of data for the stat's parameter value to save in the Character Save file. This value is always treated as a signed integer. This is only used if the \"Saved\" field is enabled.";
            itemstatcost["fCallback"] = "Boolean - If equals 1, then any changes to the stat will call the Callback function which will update thgecharacter's States, skills, or item events based on the changed stat value. If equals 0, then ignore this.";
            itemstatcost["fMin"] = "Boolean - If equals 1, then the stat will have a minimum value that cannot be reduced further than that value (\"MinAccr\" field). If equals 0, then ignore this.";
            itemstatcost["MinAccr"] = "Number - The minimum value of a stat. This is only used if the \"fMin\" field is enabled.";
            itemstatcost["Encode"] = "Number - Controls how the stat will modify an item's buy, sell, and repair costs. This field uses a code value to select a function to handle the calculations. This field relies on the \"Add\", \"Multiply\" and \"ValShift\" fields. The baseline Stat Value is first modified using the \"ValShift\" field to shift the bits. This Stat Value is then used in the calculations by one of the selected functions.";
            itemstatcost["Add"] = "Number - Used as a possible parameter value for the \"Encode\" function. Flat integer modification to the Unique item's buy, sell, and repair costs. This is added after the \"Multiply\" field has modified the costs.";
            itemstatcost["Multiply"] = "Number - Used as a possible parameter value for the \"Encode\" function. Multiplicative modifier for the item's buy, sell, and repair costs. The way this value is used depends on the Encode function selected.";
            itemstatcost["ValShift"] = "Number - Used to shift the stat's input value by a number of bits to obtain the actual value when performing calculations (such as for the \"Encode\" function).";
            itemstatcost["1.09-Save Bits"] = "Number - Controls how many bits of data are allocated for the overall size of the stat when saving/reading an item from a Character Save. This value can be treated as a signed or unsigned integer, depending on the stat. This field is only used for items saved in a game version of Patch 1.09d or older";
            itemstatcost["1.09-Save Add"] = "Number - Controls how many bits of data are allocated for the stat's value when saving/reading an item from a Character Save. This value is treated as a signed integer. This field is only used for items saved in a game version of Patch 1.09d or older";
            itemstatcost["Save Bits"] = "Number - Controls how many bits of data are allocated for the overall size of the stat when saving/reading an item from a Character Save. This value can be treated as a signed or unsigned integer, depending on the stat.";
            itemstatcost["Save Add"] = "Number - Controls how many bits of data are allocated for the stat's value when saving/reading an item from a Character Save. This value is treated as a signed integer.";
            itemstatcost["Save Param Bits"] = "Number - Controls how many bits of data for the stat's parameter value to use when saving/reading an item from a Character Save. This value is always treated asan unsigned integer.";
            itemstatcost["keepzero"] = "Boolean - If equals 1, then this stat will remain on the stat change list, when being updated, even if that stat value is 0. If equals 0, then ignore this.";
            itemstatcost["op"] = "Number - This is the stat operator, used for advanced stat modification when calculating the value of a stat. This can involves using this stat and its value to modify another stat's value. This use a function ID to determine what to calculate.";
            itemstatcost["op param"] = "Number - Used as a possible parameter value for the \"op\" function.";
            itemstatcost["op base"] = "Number - Used as a possible parameter value for the \"op\" function.";
            itemstatcost["op stat1"] = "Number - Used as a possible parameter value for the \"op\" function.";
            itemstatcost["direct"] = "Boolean - If equals 1, then when the stat is being updated in certain skill functions having to do with state changes, the stat will update in relation to its \"maxstat\" field to ensure that it never exceeds that value. If equals 0, then ignore this, and the stat will simply update in these cases. This only applies to skills that use skill server function 65, 66, 81, and 82.";
            itemstatcost["maxstat"] = "Open - Controls which stat is associated with this stat to be treated as the maximum version of this stat. This means that 2 stats are essentially linked so that there can be a current version of the stat and a maximum version to control the cap of stat's value. This is used for Life, Mana, Stamina, and Durability. This field relies on the \"direct\" field to be enabled unless it is being used for the healing potion item spell.";
            itemstatcost["damagerelated"] = "Boolean - If equals 1, then this stat will be exclusive to the item and will not add to the unit. If equals 0, then ignore this, and the stat will always add to the unit. This is typically used for weapons and is important when dual wielding weapons so that when a unit attacks, then one weapon's stats do not stack with another weapon's stats.";
            itemstatcost["itemevent1"] = "Number - Uses an event that will activate the specified function defined by \"itemeventfunc#\". Referenced from the event column in Events.txt.";
            itemstatcost["itemeventfunc1"] = "Number - Specifies the function to use after the related item event occurred. Functions are defined by a numeric ID code. This is only applied based on the \"itemevent#\" field definition. Referenced by the Code value of the Event Functions Table.";
            itemstatcost["descpriority"] = "Number - Controls how this stat is sorted in item tooltips. This field is compared to the same field on other stats to determine how to order the stats. The higher the value means that the stat will be sorted higher than other stats. If more than 1 stat has the same \"descpriority\" value, then they will be listed in the order defined in this data file.";
            itemstatcost["descfunc"] = "Number - Controls how the stat is displayed in tooltips. Uses the Code value from the below table to format the string value using the specified function.";
            itemstatcost["descval"] = "Number - Used as a possible parameter value for the descfunc function. This controls the how the value of the stat is displayed.";
            itemstatcost["descstrpos"] = "Open - Used as a possible parameter value for the descfunc function. This uses a string to display the item stat in a tooltip when its value is positive.";
            itemstatcost["descstrneg"] = "Open - Used as a possible parameter value for the descfunc function. This uses a string to display the item stat in a tooltip when its value is negative.";
            itemstatcost["descstr2"] = "Open - Used as a possible parameter value for the descfunc function. This uses a string to append to an item stat's string in a tooltip.";
            itemstatcost["dgrp"] = "Number - Assigns the stat to a group ID value. If all stats with a matching \"dgrp\" value are applied on the unit, then instead of displaying each stat individually, the group description will be applied instead (dgrpfunc field below).";
            itemstatcost["dgrpfunc"] = "Number - Controls how the shared group of stats is displayed in tooltips. Uses an ID value to select a description function to format the string value. This function IDs are exactly the same as the descfunc field.";
            itemstatcost["dgrpval"] = "Number - Used as a possible parameter value for the dgrpfunc function. This controls the how the value of the stat is displayed. (Functions the same as the \"descval\" field).";
            itemstatcost["dgrpstrpos"] = "Open - Used as a possible parameter value for the dgrpfunc function. This uses a string to display the item stat in a tooltip when its value is positive.";
            itemstatcost["dgrpstrneg"] = "Open - Used as a possible parameter value for the dgrpfunc function. This uses a string to display the item stat in a tooltip when its value is negative.";
            itemstatcost["dgrpstr2"] = "Open - Used as a possible parameter value for the dgrpfunc function. This uses a string to append to an item stat's string in a tooltip.";
            itemstatcost["stuff"] = "Number - Used as a bit shift value for handling the conversion of skill IDs and skill levels to bit values for the stat. Controls the numeric range of possible skill IDs and skill levels for charge based items. This value cannot be less than or equal to 0, or greater than 8, otherwise it will default to 6. The row that this value appears in the data file is unrelated, since this is a universally applied value.";
            itemstatcost["advdisplay"] = "Number - Controls how the stat appears in the Advanced Stats UI.";

            itemtypes["ItemType"] = "Open - This is a reference field to define the Item Type name.";
            itemtypes["Code"] = "Open - Defines the unique pointer for this Item Type, which is used by various data files.";
            itemtypes["Equiv1"] = "Open - Points to the index of another Item Type to reference as a parent. This is used to create a hierarchy for Item Types where the parents will have more universal settings shared across the related children.";
            itemtypes["Repair"] = "Boolean - If equals 1, then the item can be repaired by an NPC in the shop UI. If equals 0, then the item cannot be repaired.";
            itemtypes["Body"] = "Boolean - If equals 1, then the item can be equipped by a character (Also uses the BodyLoc1 field(s) as parameters). If equals 0, then the item can only be carried in the inventory, stash, or Horadric Cube.";
            itemtypes["BodyLoc1"] = "Open - These are required parameters if the Body field is enabled. These fields specify the inventory slots where the item can be equipped. Referenced from the Code column in BodyLocs.txt.";
            itemtypes["Shoots"] = "Number - Points to the index of another Item Type as the required equipped Item Type to be used as ammo.";
            itemtypes["Quiver"] = "Number - Points to the index of another Item Type as the required equipped Item Type to be used as this ammo's weapon.";
            itemtypes["Throwable"] = "Boolean - If equals 1, then it determines that this item is a throwing weapon. If equals 0, then ignore this.";
            itemtypes["Reload"] = "Boolean - If equals 1, then the item (considered ammo in this case) will be automatically transferred from the inventory to the required BodyLoc1 when another item runs out of that specific ammo. If equals 0, then ignore this.";
            itemtypes["ReEquip"] = "Boolean - If equals 1, then the item in the inventory will replace a matching equipped item if that equipped item was destroyed. If equals 0, then ignore this.";
            itemtypes["AutoStack"] = "Boolean - If equals 1, then if the player picks up a matching Item Type, then they will try to automatically stack together. If equals 0, then ignore this.";
            itemtypes["Magic"] = "Boolean - If equals 1, then this item will always have the Magic quality (unless it is a Quest item). If equals 0, then ignore this.";
            itemtypes["Rare"] = "Boolean - If equals 1, then this item can spawn as a Rare quality. If equals 0, then ignore this.";
            itemtypes["Normal"] = "Boolean - If equals 1, then this item will always have the Normal quality. If equals 0, then ignore this.";
            itemtypes["Beltable"] = "Boolean - If equals 1, then this item can be placed in the character's belt slots. If equals 0, then ignore this.";
            itemtypes["MaxSockets1"] = "Number - Determines the maximum possible number of sockets that can be spawned on the item when the item level is greater than or equal to 1 and less than or equal to the matching MaxSocketsLevelThreshold1 value(s). The number of sockets is also determined by the Gemsockets value from AMW.txt.";
            itemtypes["MaxSocketsLevelThreshold1"] = "Number - Defines the item level thresholds using the above MaxSockets1 field(s).";
            itemtypes["TreasureClass"] = "Boolean - If equals 1, then allow this Item Type to be used in default treasure classes. If equals 0, then ignore this.";
            itemtypes["Rarity"] = "Number - Determines the chance for the item to spawn with stats, when created as a random Weapon/Armor/Misc item. Used in the following formula: IF RANDOM(0, ([\"Rarity\"] - [Current Act Level]))> 0, THEN spawn stats.";
            itemtypes["StaffMods"] = "Open - Determines if the Item Type should have class specific item skill modifiers. Grants skills to the class listed using the Code column in PlayerClass.txt.";
            itemtypes["Class"] = "Open - Determines if this item should be useable only by a specific class. Referenced from the Code column in PlayerClass.txt.";
            itemtypes["VarInvGfx"] = "Number - Tracks the number of inventory graphics used for this item type. This number much match the number of InvGfx1 field(s) used.";
            itemtypes["InvGfx1"] = "Open - Defines a DC6 file to use for the item's inventory graphics. The entry amount should equal the value used in above VarInvGfx.";
            itemtypes["StorePage"] = "Open - Uses a code to determine which UI tab page on the NPC shop UI to display this Item Type, such as after it is sold to the NPC. Referenced from the Code column in StorePage.txt.";

            levelgroups["Name"] = "Open - Defines the unique name pointer for the level group, which is used in other files.";
            levelgroups["ID"] = "Number - Defines the unique numeric ID for the level group, which is used in other files.";
            levelgroups["GroupName"] = "Open - String Field. Used for displaying the name of the level group, such as when all levels in a group have been desecrated (Terrorized).";

            levels["Name"] = "Open - Defines the unique name pointer for the area level, which is used in other files.";
            levels["ID"] = "Number - Defines the unique numeric ID for the area level, which is used in other files.";
            levels["Pal"] = "Number - Defines which palette file to use for the area level. This uses index values from 0 to 4 to convey Act 1 to Act 5.";
            levels["Act"] = "Number - Defines the Act number that the area level is a part of. This uses index values from 0 to 4 to convey Act 1 to Act 5.";
            levels["QuestFlag"] = "Number - Controls what quest record that the player needs to have completed before being allowed to enter this area level, while playing in Classic Mode. Each quest can have multiple quest records, and this field is looking for a specific quest record from a quest. Referenced by the Code value of the Quest Flags Table.";
            levels["QuestFlagEx"] = "Number - Controls what quest record that the player needs to have completed before being allowed to enter this area level, while playing in Expansion Mode. Each quest can have multiple quest records, and this field is looking for a specific quest record from a quest. Referenced by the Code value of the Quest Flags Table.";
            levels["Layer"] = "Number - Defines a unique numeric ID that is used to identify which Automap data belongs to which area level when saving and loading data from the character save.";
            levels["SizeX"] = "Number - Specifies the Length tile size values of an entire area level, which are used for determining how to build the level, for Normal, Nightmare, and Hell Difficulty, respectively.";
            levels["SizeY"] = "Number - Specifies the Width tile size values of an entire area level, which are used for determining how to build the level, for Normal, Nightmare, and Hell Difficulty, respectively.";
            levels["OffsetX & OffsetY"] = "Number - Specifies the location offset coordinates (measured in tile size) for the origin point of the area level in the world. Must not overlap with other level offsets";
            levels["Depend"] = "Number - Assigns another level to be this area level's depended level, which controls this area level's position and how it starts building its tiles. Uses the level ID field. If this equals 0, then ignore this.";
            levels["Teleport"] = "Number - Controls the functionality of the Sorceress Teleport skill and the Assassin Dragon Flight skill on the area level.";
            levels["Rain"] = "Boolean - If equals 1, then allow rain to play its effects on the area level. If the level is part of Act 5, then it will snow on the area level, instead of rain. If equals 0, then it will never rain on the area level.";
            levels["Mud"] = "Boolean - If equals 1, then random bubbles will animate on the tiles that are flagged as water tiles. If equals 0, then ignore this.";
            levels["NoPer"] = "Boolean - If equals 1, then allow the use of display option of Perspective Mode while the player is in the level. If equals 0, then disable the option of Perspective Mode and force the player to use Orthographic Mode while the player is in the level.";
            levels["LOSDraw"] = "Boolean - If equals 1, then the level will check the player's line of sight before drawing monsters. If equals 0, then ignore this.";
            levels["FloorFilter"] = "Boolean - If equals 1 and if the floor's layer in the area level equals 1, then draw the floor tiles with a linear texture sampler. If equals 0, then draw the floor tiles with a nearest texture sampler.";
            levels["BlankScreen"] = "Boolean - If equals 1, then draw the area level screen. If equals 0, then do not draw the area level screen, meaning that the level will be a blank screen.";
            levels["DrawEdges"] = "Boolean - If equals 1, then draw the areas in levels that are not covered by floor tiles. If equals 0, then ignore this.";
            levels["DrlgType"] = "Number - Determines the type of Dynamic Random Level Generation used for building and handling different elements of the area level. Uses the Code from the table below to handle which type of DRLG is used.";
            levels["LevelType"] = "Number - Defines the Level Type used for this area level. Uses the ID field from LvlTypes.txt.";
            levels["SubType"] = "Number - Controls the group of tile substitutions for the area level (LvlSub.txt). There are defined sub-types to choose from, listed below.";
            levels["SubTheme"] = "Number - Controls which theme number to use in a Level Substitution. The allowed values are 0 to 4, which convey which Prob0, Trials0, and Max0 field to use from the LvlSub.txt file. If this equals -1, then there is no sub theme for the area level.";
            levels["SubWaypoint"] = "Number - Controls the level substitutions for adding waypoints in the area level (LvlSub.txt). This uses a defined sub type to choose from and also depends on the room having a waypoint tile.";
            levels["SubShrine"] = "Number - Controls the level substitutions for adding Shrines in the area level (LvlSub.txt). This uses a defined sub type to choose from and also depends on the room allowing for a shrine to spawn.";
            levels["Vis0"] = "Number - Defines the visibility of other area levels involved with this area level, allowing for travel functionalities between levels. This uses the ID field of another defined area level to link with this area level. If this equals 0, then no area level is specified.";
            levels["Warp0"] = "Number - Uses the ID field from LvlWarp.txt, which defines which Level Warp to use when exiting the area level. This is connected with the definition of the related Vis0 field. If this equals -1, then no Level Warp is specified which should also mean that the related Vis0 field is not defined.";
            levels["Intensity"] = "Number - Controls the intensity value of the area level's ambient colors. This affects brightness of the room's RGB colors. Uses a value between 0 and 255. If all these related fields equal 0, then the game ignores setting the area level's ambient colors.";
            levels["Red"] = "Number - Controls the red value of the area level's ambient colors. Uses a value between 0 and 255.";
            levels["Green"] = "Number - Controls the green value of the area level's ambient colors. Uses a value between 0 and 255.";
            levels["Blue"] = "Number - Controls the blue value of the area level's ambient colors. Uses a value between 0 and 255.";
            levels["Portal"] = "Boolean - If equals 1, then this area level will be flagged as a portal level, which is saved in the player's information and can be used for keeping track of the player's portal functionalities. If equals 0, then ignore this.";
            levels["Position"] = "Boolean - If equals 1, then enable special casing for positioning the player on the area level. This can mean that the player could spawn on a different location on the area level, depending on the level room's position type. An example can be when the player spawns in a town when loading the game, or using a waypoint, or using a town portal. If equals 0, then ignore this.";
            levels["SaveMonsters"] = "Boolean - If equals 1, then the game will save the monsters in the area level, such as when all players leave the area level. If equals 0, then monsters will not be saved and will be removed. This is usually disabled for areas where monsters do not spawn.";
            levels["Quest"] = "Number - Controls what quest record is attached to monsters that spawn in this area level. This is used for specific quests handling lists of monsters in the area level. Referenced by the Code value of the Quest Flags Table.";
            levels["WarpDist"] = "Number - Defines the minimum pixel distance from a Level Warp that a monster is allowed to spawn near. Tile distance values are converted to game pixel distance values by multiplying the tile distance value by 160 / 32, where 160 is the width of pixels of a tile.";
            levels["MonLvl"] = "Number - Controls the overall monster level for the area level for Normal, Nightmare, and Hell Difficulty, respectively. This is for Classic mode only. This can affect the highest item level allowed to drop in this area level.";
            levels["MonLvlEx"] = "Number - Controls the overall monster level for the area level for Normal, Nightmare, and Hell Difficulty, respectively. This is for Expansion mode only. This can affect the highest item level allowed to drop in this area level.";
            levels["MonDen"] = "Number - Controls the monster density on the area level for Normal, Nightmare, and Hell Difficulty, respectively. This is a random value out of 100000, which will determine whether to spawn or not spawn a monster pack in the room of the area level. If this value equals 0, then no random monsters will populate on the area level.";
            levels["MonUMin"] = "Number - Defines the minimum number of Unique Monsters that can spawn in the area level for Normal, Nightmare, and Hell Difficulty, respectively. This field depends on the related MonDen fields being defined.";
            levels["MonUMax"] = "Number - Defines the maximum number of Unique Monsters that can spawn in the area level for Normal, Nightmare, and Hell Difficulty, respectively. This field depends on the related MonDen field being defined. Each room in the area level will attempt to spawn a Unique Monster with a 5/100 random chance, and this field's value will cap the number of successful attempts for the entire area level.";
            levels["MonWndr"] = "Boolean - If equals 1, then allow Wandering Monsters to spawn on this area level (WanderingMon.txt). This field depends on the related MonDen field being defined. If equals 0, then ignore this.";
            levels["MonSpcWalk"] = "Number - Defines a distance value, used to handle monster pathing AI when the level has certain pathing blockers, such as jail bars or rivers. In these cases, monsters will walk randomly until a player is located within this distance value or when the monsters find a possible path to target the player. If this equals 0, then ignore this field.";
            levels["NumMon"] = "Number - Controls the number of different monsters randomly spawn in the area level. The maximum value is 13. This controls the number of random selections from the 25 related mon1 and nmon1 fields or umon1 fields, depending on the game difficulty and monster type.";
            levels["mon1"] = "Open - Defines which monsters can spawn on the area level for Normal Difficulty. Uses the monster ID field from MonStats.txt.";
            levels["rangedspawn"] = "Boolean - If equals 1, then for the first monster, try to pick a ranged type. If equals 0, then ignore this.";
            levels["nmon1"] = "Open - Defines which monsters can spawn on the area level for Nightmare Difficulty and Hell Difficulty. Uses the monster ID field from MonStats.txt.";
            levels["umon1"] = "Open - Defines which monsters can spawn as Unique monsters on this area level for Normal Difficulty. Uses the monster ID field from MonStats.txt.";
            levels["cmon1"] = "Open - Defines which Critter monsters can spawn on the area level. Uses the monster ID field from MonStats.txtCritter monsters are determined by the critter field from MonStats2.txt.";
            levels["cpct1"] = "Number - Controls the percent chance (out of 100) to spawn a Critter monster on the area level.";
            levels["camt1"] = "Number - Controls the amount of Critter monsters to spawn on the area level after they succeeded their random spawn chance from the above cpct1 field.";
            levels["Themes"] = "Number - Controls the type of theme when building a level room. This value is a summation of possible values to build a bit mask for determining which themes to use when building a level room. For example, a value of 60 means that the area level can have the following themes: 32, 16, 8, 4.";
            levels["SoundEnv"] = "Number - Uses the Index field from SoundEnviron.txt, which controls what music is played while the player is in the area level.";
            levels["Waypoint"] = "Number - Defines the unique numeric ID for the Waypoint in the area level. If this value is greater than or equal to 255, then ignore this field.";
            levels["LevelName"] = "Open - Used for displaying the name of the area level, such as when in the UI when the Automap is being viewed.";
            levels["LevelWarp"] = "Open - Used for displaying the entrance name of the area level on Level Warp tiles that link to this area level. For example, when the player mouse hovers over a tile to warp to the area level, then this string is displayed.";
            levels["LevelEntry"] = "Open - Used for displaying the UI popup title string when the player enters the area level.";
            levels["ObjGrp0"] = "Number - Uses a numeric *ID Index to define which possible Object Groups to spawn in this area level (ObjGroup.txt). The game will go through each of these fields, so there can be more than 1 Object Group used in an area level. If this value equals 0, then ignore this.";
            levels["ObjPrb0"] = "Number - Determines the random chance (out of 100) for each Object Group to spawn in the area level. This field depends on the related \"ObjGrp#\" field being defined.";
            levels["LevelGroup"] = "Open - Defines what group this level belongs to. Used for condensing level names in desecrated (terror) zones messaging. See LevelGroups.txt.";

            lvlmaze["Name"] = "Open - This is a reference field to describe the area level. IDeally this should match the name of the area level from the Levels.txt file.";
            lvlmaze["Level"] = "Number - This refers to the ID field from the Levels.txt file.";
            lvlmaze["Rooms"] = "Number - Controls the total number of rooms that a Level Maze will generate when playing the game in Normal Difficulty, Nightmare Difficulty, and Hell Difficulty, respectively.";
            lvlmaze["Merge"] = "Number - This value affects the probability that a room gets an adjacent room linked next to it. This is a random chance out of 1000.";

            lvlprest["Name"] = "Open - This is a reference field to define the Level Preset.";
            lvlprest["Def"] = "Number - Defines the unique numeric ID for the Level Preset. This is referenced in other files.";
            lvlprest["LevelID"] = "Number - This refers to the ID field from the Levels.txt file. If this value is not equal to 0, then this Level Preset is used to build that entire area level. If this value is equal to 0, then the Level Preset does not define the entire area level and is used as a part of constructing area levels.";
            lvlprest["Populate"] = "Boolean - If equals 1, then units are allowed to spawn in the Level Preset. If equals 0, then units will never spawn in the Level Preset.";
            lvlprest["Logicals"] = "Boolean - If equals 1, then the Level Preset allow for wall transparency to function. If equals 0, then walls will always appear solid.";
            lvlprest["Outdoors"] = "Boolean - If equals 1, then the Level Preset will be classified as an outdoor area, which can mean that lighting will function differently. If equals 0, then the Level Preset will be classified as an indoor area.";
            lvlprest["Animate"] = "Boolean - If equals 1, then the game will animate the tiles in the Level Preset. If equals 0, then ignore this.";
            lvlprest["KillEdge"] = "Boolean - If equals 1, then the game will remove tiles that border the size of the Level Preset. If equals 0, then ignore this.";
            lvlprest["FillBlanks"] = "Boolean - If equals 1, then all blank tiles in the Level Preset will be filled with unwalkable tiles. If equals 0, then ignore this.";
            lvlprest["AutoMap"] = "Boolean - If equals 1, then this Level Preset will be automatically completely revealed on the Automap. If equals 0, then this Level Preset will be hidden on the Automap and will need to be explored.";
            lvlprest["Scan"] = "Boolean - If equals 1, then this Level Preset will allow the usage of warping with waypoints (This requires that the Level Preset has a waypoint object). If equals 0, then ignore this.";
            lvlprest["Pops"] = "Number - Defines how many Pop tiles are defined in the Level Preset file. These Pop tiles are mainly used for controlling the roof and wall popping when a player enters a building in an area.";
            lvlprest["PopPad"] = "Number - Determines the size of the Pop tile area, by using an offset value. This offset value can increase or decrease the size of the Pop tile size if it has a positive or negative value.";
            lvlprest["Files"] = "Number - Determines the number of different versions to use for the Level Preset. This value acts as a range, which the game will use for randomly choosing one of the File1 fields to build the Level Preset. This is how the Level Presets have variety when the area level is being built.";
            lvlprest["File1"] = "Open - Specifies the name of which ds1 file to use. The ds1 files contain data for building Level Presets. If this value equals 0, then this field will be ignored. The number of these defined fields should match the value used in the \"Files\" field.";
            lvlprest["Dt1Mask"] = "Number - This functions as a bit field mask with a size of a 32 bit value. This explains to the ds1 file which of the 32 dt1 tile files to use from a Level Type when assembling the Level Preset. Each File 1 field value from LevelTypes.txt is assigned a bit value, up to the 32 possible bit values. (For example: File1 = 1, File2=2, File3 = 4, File4=8, File5=16….File32 = 2147483648). To build the \"Dt1Mask\", you would add their associated bit values together for a total value. This total value is the bitmask value.";

            lvlsub["Name"] = "Open - This is a reference field to describe the Level Substitution.";
            lvlsub["Type"] = "Number - This refers to the SubType field from the Levels.txt file. This defines a group that multiple substitutions can share.";
            lvlsub["File"] = "Open - Specifies the name of which ds1 file to use. The ds1 files contain data for building Level Presets.";
            lvlsub["CheckAll"] = "Boolean - If equals 1, then substitute each tile in the room. If equals 0, then substitute random tiles in the room.";
            lvlsub["BordType"] = "Number - This controls how often substituting tiles can work for border tiles.";
            lvlsub["GridSize"] = "Number - Controls the tile size of a cluster for substituting tiles. This evenly affects both the X and Y size values of a room.";
            lvlsub["Dt1Mask"] = "Number - This functions as a bit field mask with a size of a 32 bit value. This explains to the ds1 file which of the 32 dt1 tile files to use from a Level Type when assembling the Level Preset. Each File 1 field value from LevelTypes.txt is assigned a bit value, up to the 32 possible bit values. (For example: File1 = 1, File2=2, File3 = 4, File4=8, File5=16….File32 = 2147483648). To build the \"Dt1Mask\", you would add their associated bit values together for a total value. This total value is the bitmask value.";
            lvlsub["Prob0"] = "Number - This value affects the probability that the tile substitution is used. This is a random chance out of 100. Which \"Prob#\" field that is checked depends on the SubTheme value from the Levels.txt file.";
            lvlsub["Trials0"] = "Number - Controls the number of times to randomly substitute tiles in a cluster. If this value equals -1, then the game will try to do as many tile substitutions that can be allowed based on the cluster and tile size. This field depends on the CheckAll field being equal to 0.";
            lvlsub["Max0"] = "Number - The maximum number of clusters of tiles to substitute randomly. This field depends on the CheckAll field being equal to 0.";

            lvltypes["Name"] = "Open - This is a reference field to define the name of each Level Type.";
            lvltypes["ID"] = "Number - This is a reference field to define the index of each Level Type.";
            lvltypes["File 1"] = "Open - Specifies the name of which dt1 file to use. The dt1 files contain the images for each area tile found in each Act. If this value equals 0, then this field will be ignored.";
            lvltypes["Act"] = "Number - Defines which Act is related to the Level Type. When loading an Act, the game will only use the Level Types associated with that Act number. Uses a decimal number to convey each Act number (Ex: A value of 3 means Act 3).";

            lvlwarp["Name"] = "Open - This is a reference field to define the Level Warp.";
            lvlwarp["ID"] = "Number - Defines the numeric ID for the type of Level Warp. This ID can be shared between multiple Level Warps if those Level Warps want to use the same functionality. This is referenced in other files";
            lvlwarp["SelectX & SelectY"] = "Number - If equals 1, then Level Warp tiles will change their appearance when highlighted. If equals 0, then the Level Warp tiles will not change appearance when highlighted.";
            lvlwarp["SelectDX & SelectDY"] = "Number - If equals 1, then Level Warp tiles will change their appearance when highlighted. If equals 0, then the Level Warp tiles will not change appearance when highlighted.";
            lvlwarp["ExitWalkX & ExitWalkY"] = "Number - If equals 1, then Level Warp tiles will change their appearance when highlighted. If equals 0, then the Level Warp tiles will not change appearance when highlighted.";
            lvlwarp["OffsetX & OffsetY"] = "Number - If equals 1, then Level Warp tiles will change their appearance when highlighted. If equals 0, then the Level Warp tiles will not change appearance when highlighted.";
            lvlwarp["LitVersion"] = "Boolean - If equals 1, then Level Warp tiles will change their appearance when highlighted. If equals 0, then the Level Warp tiles will not change appearance when highlighted.";
            lvlwarp["Tiles"] = "Number - Defines an index offset to determine which tile to use in the tile set for the highlighted version of the Level Warp. These tiles are loaded and hidden/revealed when the player mouse hovers over the Level Warp tiles. This relies on LitVersion being enabled.";
            lvlwarp["NoInteract"] = "Boolean - If equals 1, then the Level War cannot be directly interacted by the player. If equals 0, then the player can interact with the Level Warp.";
            lvlwarp["Direction"] = "Open - Defines the orientation of the Level Warp. Uses a specific string code.";
            lvlwarp["UniqueID"] = "Number - Defines the unique numeric ID for the Level Warp. Each Level Warp should have a unique ID so that the game can handle loading that specific Level Warp's related files.";

            magicprefix["Name"] = "Open - Defines the item affix name.";
            magicprefix["version"] = "Number - Defines which game version to use this item affix (<100 = Classic mode | 100 = Expansion mode).";
            magicprefix["spawnable"] = "Boolean - If equals 1, then this item affix is used as part of the game's randomizer for assigning item modifiers when an item spawns. If equals 0, then this item affix is never used.";
            magicprefix["rare"] = "Boolean - If equals 1, then this item affix can be used when randomly assigning item modifiers when a rare item spawns. If equals 0, then this item affix is not used for rare items.";
            magicprefix["level"] = "Number - The minimum item level required for this item affix to spawn on the item. If the item level is below this value, then the item affix will not spawn on the item.";
            magicprefix["maxlevel"] = "Number - The maximum item level required for this item affix to spawn on the item. If the item level is above this value, then the item affix will not spawn on the item.";
            magicprefix["levelreq"] = "Number - The minimum character level required to equip an item that has this item affix.";
            magicprefix["classspecific"] = "Open - Controls if this item affix should only be used for class specific items. This relies on the Class specified from ItemTypes.txt, for the specific item. Referenced from the Code column in PlayerClass.txt.";
            magicprefix["class"] = "Open - Controls which character class is required for the class specific level requirement (classlevelreq field). Referenced from the Code column in PlayerClass.txt.";
            magicprefix["classlevelreq"] = "Number - The minimum character level required for a specific class in order to equip an item that has this item affix. This relies on the class specified. If equals null, then the class will default to using the levelreq field.";
            magicprefix["frequency"] = "Number - Controls the probability that the affix appears on the item (a higher value means that the item affix will appear on the item more often). This value gets summed together with other \"frequency\" values from all possible item affixes that can spawn on the item, and then is used as a denominator value for the randomizer. Whichever item affix is randomly selected will be the one to appear on the item. The formula is calculated as the following: [Item Affix Selected] = [\"frequency\"] / [Total Frequency]. If the item has a magic level (from the \"magic lvl\" field in Weapons.txt/Armor.txt/Misc.txt) then the magic level value is multiplied with this value. If equals 0, then this item affix will never appear on an item.";
            magicprefix["group"] = "Number - Assigns an item affix to a specific group number. Items cannot spawn with more than 1 item affix with the same group number. This is used to guarantee that certain item affixes do not overlap on the same item. If this field is null, then the group number will default to group 0.";
            magicprefix["mod1code"] = "Open - Controls the item properties for the item affix (Uses the Code field from Properties.txt).";
            magicprefix["mod1param"] = "Number - The \"parameter\" value associated with the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicprefix["mod1min"] = "Number - The \"min\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicprefix["mod1max"] = "Number - The \"max\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicprefix["transformcolor"] = "Open - Controls the color change of the item after spawning with this item affix. If empty, then the item affix will not change the item's color. Referenced from the Code column in Colors.txt.";
            magicprefix["itype1"] = "Open - Controls what Item Types are allowed to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            magicprefix["etype1"] = "Open - Controls what Item Types are excluded to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            magicprefix["multiply"] = "Number - Multiplicative modifier for the item's buy and sell costs, based on the item affix (Calculated in 1024ths for buy cost and 4096ths for sell cost).";
            magicprefix["add"] = "Number - Flat integer modification to the item's buy and sell costs, based on the item affix.";

            magicsuffix["Name"] = "Open - Defines the item affix name.";
            magicsuffix["version"] = "Number - Defines which game version to use this item affix (<100 = Classic mode | 100 = Expansion mode).";
            magicsuffix["spawnable"] = "Boolean - If equals 1, then this item affix is used as part of the game's randomizer for assigning item modifiers when an item spawns. If equals 0, then this item affix is never used.";
            magicsuffix["rare"] = "Boolean - If equals 1, then this item affix can be used when randomly assigning item modifiers when a rare item spawns. If equals 0, then this item affix is not used for rare items.";
            magicsuffix["level"] = "Number - The minimum item level required for this item affix to spawn on the item. If the item level is below this value, then the item affix will not spawn on the item.";
            magicsuffix["maxlevel"] = "Number - The maximum item level required for this item affix to spawn on the item. If the item level is above this value, then the item affix will not spawn on the item.";
            magicsuffix["levelreq"] = "Number - The minimum character level required to equip an item that has this item affix.";
            magicsuffix["classspecific"] = "Open - Controls if this item affix should only be used for class specific items. This relies on the Class specified from ItemTypes.txt, for the specific item. Referenced from the Code column in PlayerClass.txt.";
            magicsuffix["class"] = "Open - Controls which character class is required for the class specific level requirement (classlevelreq field). Referenced from the Code column in PlayerClass.txt.";
            magicsuffix["classlevelreq"] = "Number - The minimum character level required for a specific class in order to equip an item that has this item affix. This relies on the class specified. If equals null, then the class will default to using the levelreq field.";
            magicsuffix["frequency"] = "Number - Controls the probability that the affix appears on the item (a higher value means that the item affix will appear on the item more often). This value gets summed together with other \"frequency\" values from all possible item affixes that can spawn on the item, and then is used as a denominator value for the randomizer. Whichever item affix is randomly selected will be the one to appear on the item. The formula is calculated as the following: [Item Affix Selected] = [\"frequency\"] / [Total Frequency]. If the item has a magic level (from the \"magic lvl\" field in Weapons.txt/Armor.txt/Misc.txt) then the magic level value is multiplied with this value. If equals 0, then this item affix will never appear on an item.";
            magicsuffix["group"] = "Number - Assigns an item affix to a specific group number. Items cannot spawn with more than 1 item affix with the same group number. This is used to guarantee that certain item affixes do not overlap on the same item. If this field is null, then the group number will default to group 0.";
            magicsuffix["mod1code"] = "Open - Controls the item properties for the item affix (Uses the Code field from Properties.txt).";
            magicsuffix["mod1param"] = "Number - The \"parameter\" value associated with the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicsuffix["mod1min"] = "Number - The \"min\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicsuffix["mod1max"] = "Number - The \"max\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            magicsuffix["transformcolor"] = "Open - Controls the color change of the item after spawning with this item affix. If empty, then the item affix will not change the item's color. Referenced from the Code column in Colors.txt.";
            magicsuffix["itype1"] = "Open - Controls what Item Types are allowed to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            magicsuffix["etype1"] = "Open - Controls what Item Types are excluded to spawn with this item affix. Uses the Code field from ItemTypes.txt.";
            magicsuffix["multiply"] = "Number - Multiplicative modifier for the item's buy and sell costs, based on the item affix (Calculated in 1024ths for buy cost and 4096ths for sell cost).";
            magicsuffix["add"] = "Number - Flat integer modification to the item's buy and sell costs, based on the item affix.";

            missiles["Missile"] = "Open - Defines the unique name ID for the missile, which is how other files can reference the missile. The order of defined missiles will determine their ID numbers, so they should not be reordered.";
            missiles["pCltDoFunc"] = "Number - Uses an ID value from the below table, to select a function for the missile's behavior that will be active every frame on the client side. This is more about handling the local graphics while the missile is moving.";
            missiles["pCltHitFunc"] = "Number - Uses an ID value from the below table, to select a specialized function for the missile's behavior when hitting something on the client side. This is more about handling the local graphics at the moment of missile collision.";
            missiles["pSrvDoFunc"] = "Number - Uses an ID value from the below table, to select a specialized function for the missile's behavior while active every frame on the server side.";
            missiles["pSrvHitFunc"] = "Number - Uses an ID value from the below table, to select a specialized function for the missile's behavior when hitting something on the server side.";
            missiles["pSrvDmgFunc"] = "Number - Uses an ID value from the below table, to select a specialized function that gets called before damaging a unit on the server side.";
            missiles["SrvCalc1"] = "Calculation - Used as a parameter for the pSrvDoFunc field.";
            missiles["Param1"] = "Number - Used as a parameter for the pSrvDoFunc field.";
            missiles["CltCalc1"] = "Calculation - Used as a parameter for the pCltDoFunc field.";
            missiles["CltParam1"] = "Number - Used as a parameter for the pCltDoFunc field.";
            missiles["SHitCalc1"] = "Calculation - Used as a parameter for the pSrvHitFunc field.";
            missiles["sHitPar1"] = "Number - Used as a parameter for the pSrvHitFunc field.";
            missiles["CHitCalc1"] = "Calculation - Used as a parameter for the pCltHitFunc field.";
            missiles["cHitPar1"] = "Number - Used as a parameter for the pCltHitFunc field.";
            missiles["DmgCalc1"] = "Calculation - Used as a parameter for the pSrvDmgFunc field.";
            missiles["dParam1"] = "Number - Used as a parameter for the pSrvDmgFunc field.";
            missiles["Vel"] = "Number - The baseline velocity of the missile, which is the speed at which the missile moves in the game world. This is measured by distance in pixels traveled per frame.";
            missiles["MaxVel"] = "Number - The maximum velocity of the missile. If the missile's current velocity increases (based on other fields), then this field controls how high the velocity is allowed to go.";
            missiles["VelLev"] = "Number - Adds extra velocity based on the caster unit's level. Each level gained beyond level 1 will add this value to the baseline Vel field.";
            missiles["Accel"] = "Number - Controls the acceleration of the missile's movement. A positive value will increase the missile's velocity per frame. A negative value will decrease the missile's velocity per frame. The bigger positive or negative values will cause the velocity to change faster per frame.";
            missiles["Range"] = "Number - Controls the baseline duration that the missile will exist for after it is created. This is measured in frames where 25 Frames = 1 second.";
            missiles["LevRange"] = "Number - Adds extra duration based on the caster unit's level. Each level gained beyond level 1 will add this value to the baseline Range field.";
            missiles["Light"] = "Number - Controls the missile's Light Radius size (measured in grid sub-tiles).";
            missiles["Flicker"] = "Number - If greater than 0, then every 4th frame while the missile is active, the Light Radius will randomly change in size between base size to its base size plus this value (measured in grid sub-tiles).";
            missiles["Red"] = "Number - Controls the red color value of the missile's Light Radius (Uses a value from 0 to 255).";
            missiles["Green"] = "Number - Controls the green color value of the monster's Light Radius (Uses a value from 0 to 255).";
            missiles["Blue"] = "Number - Controls the blue color value of the monster's Light Radius (Uses a value from 0 to 255).";
            missiles["InitSteps"] = "Number - The number of frames the missile needs to be alive until it becomes visible on the game client. If the missile's current duration in frame count is less than this value, then the missile will appear invisible.";
            missiles["Activate"] = "Number - The number of frames the missile needs to be alive until it becomes active. If the missile's current duration in frame count is less than this value, then the missile will not collide.";
            missiles["LoopAnim"] = "Boolean - If equals 1, then the missile's animation will repeat once the previous animation finishes. If equals 0, then the missile's animation will only play once, which can cause the missile to appear invisible at the end of the animation, but it will still be alive.";
            missiles["CelFile"] = "Open - Defines which DCC missile file to use for the visual graphics of the missile.";
            missiles["animrate"] = "Number - Controls the visual speed of the missile's animation graphics. The overall missile animation rate is calculated as the following: 256 * [\"animrate\"] / 1024.";
            missiles["AnimLen"] = "Number - Defines the length of the missile's animation in frames where 25 Frames = 1 second. This field can sometimes be used to calculate the missile animation rate, depending on the missile function used.";
            missiles["AnimSpeed"] = "Number - Controls the visual speed of the missile's animation graphics on the client side (Measured in 16ths, where 16 equals 1 frame per second). This can be overridden by certain missile functions.";
            missiles["RandStart"] = "Number - If this value is greater than 0, then the missile will start at a random frame between 0 and this value when it begins its animation.";
            missiles["SubLoop"] = "Boolean - If equals 1, then the missile will use a specific sequence of its animation while it is alive, depending on its creation. If equals 0, then the missile will not use a sequenced animation.";
            missiles["SubStart"] = "Number - The starting frame of the sequence animation. This requires that the SubLoop field is enabled.";
            missiles["SubStop"] = "Number - The ending frame of the sequence animation. After reaching this frame, then the sequenced animation will loop back to the SubStart frame. This requires that the SubLoop field is enabled.";
            missiles["CollideType"] = "Number - Defines the missile's collision type, which controls what units, objects, or parts of the environment that the missile can impact.";
            missiles["CollideKill"] = "Boolean - If equals 1, then the missile will be destroyed when it collides with something. If equals 0, then the missile will not be destroyed when it collides with something.";
            missiles["CollideFriend"] = "Boolean - If equals 1, then the missile can collide with friendly units, including the caster. If equals 0, then the missile will ignore friendly units.";
            missiles["LastCollide"] = "Boolean - If equals 1, then the missile will track the last unit that it collided with, which is useful for making sure the missile does not hit the same unit twice. If equals 0, then ignore this.";
            missiles["Collision"] = "Boolean - If equals 1, then the missile will have a missile type path placement collision mask when it is initialized or moved. If equals 0, then the missile will have no placement collision mask when it is created or moved.";
            missiles["ClientCol"] = "Boolean - If equals 1, then the missile will check collision on the client, depending on the missile's CollideType field. If equals 0, then ignore this.";
            missiles["ClientSend"] = "Boolean - If equals 1, then the server will create the missile on the client. This can be used when reloading area levels or transitioning units between areas. If equals 0, then ignore this.";
            missiles["NextHit"] = "Boolean - If equals 1, then the missile will use the next delay. If equals 0, then ignore this.";
            missiles["NextDelay"] = "Number - Controls the delay in frame length until the missile is allowed to hit the same unit again. This field relies on the NextHit field being enabled.";
            missiles["xoffset & yoffset & zoffset"] = "Specifies the X, Y, and Z location coordinates (measured in pixels) to offset to visually draw the missile based on its actual location. This will only offset the visual graphics of the missile, not the missile itself. The Z axis controls the visual height of the missile";
            missiles["Size"] = "Number - Defines the diameter in sub-tiles (for both the X and Y axis) that the missile will occupy. This affects how the missile will collide with something or how the game will handle placement for the missile.";
            missiles["SrcTown"] = "Boolean - If equals 1, then the missile will be destroyed if the caster unit is located in an act town. If equals 0, then ignore this.";
            missiles["CltSrcTown"] = "Number - If this value is greater than 0 and the LoopAnim field is disabled, then this field will control which frame to set the missile's animation when the player is in town. This value gets subtracted from the AnimLen value to determine the frame to set the missile's animation.";
            missiles["CanDestroy"] = "Boolean - If equals 1, then the missile can be attacked and destroyed. If equals 0, then the missile cannot be attacked.";
            missiles["ToHit"] = "Boolean - If equals 1, then this missile will use the caster's Attack Rating stat to determine if the missile should hit its target. If equals 0, then the missile will always hit its target.";
            missiles["AlwaysExplode"] = "Boolean - If equals 1, then the missile will always process an explosion when it is killed, which can use the pSrvHitFunc, HitSound and ExplosionMissile fields on the client side. If equals 0, then the missile will only rely on proper collision hits to process an explosion.";
            missiles["Explosion"] = "Boolean - If equals 1, then the missile will be classified as an explosion which will make it use different handlers for finding nearby units and dealing damage. If equals 0, then ignore this.";
            missiles["Town"] = "Boolean - If equals 1, then the missile is allowed to be alive when in a town area. If equals 0, then the missile will be immediately destroyed when located within a town area.";
            missiles["NoUniqueMod"] = "Boolean - If equals 1, then the missile will not receive bonuses from Unique monster modifiers. If equals 0, then the missile will receive bonuses from Unique monster modifiers.";
            missiles["NoMultiShot"] = "Boolean - If equals 1, then the missile will not be affected by the Multi-Shot monster modifier. If equals 0, then the missile will be affected by theMulti-Shot monster modifier.";
            missiles["Holy"] = "Number - Controls a bit field flag where each value is a code to allow the missile to damage a certain type of monster.";
            missiles["CanSlow"] = "Boolean - If equals 1, then the missile can be affected by the \"slowmissiles\" state (States.txt). If equals 0, then the missile will ignore the \"slowmissiles\" state.";
            missiles["ReturnFire"] = "Boolean - If equals 1, then missile can trigger the Sorceress Chilling Armor event function. If equals 0, then this missile will not trigger that function.";
            missiles["GetHit"] = "Boolean - If equals 1, then the missile will cause the target unit to enter the Get Hit mode (GH), which acts as the hit recovery mode. If equals 0, then ignore this.";
            missiles["SoftHit"] = "Boolean - If equals 1, then the missile will cause a soft hit on the unit, which can trigger a blood splatter effect, hit flash, and/or a hit sound. If equals 0, then ignore this.";
            missiles["KnockBack"] = "Number - Controls the percentage chance (out of 100) that the target unit will be knocked back when hit by the missile.";
            missiles["Trans"] = "Number - Controls the alpha mode for how the missile is displayed, which can affect transparency and blending.";
            missiles["Pierce"] = "Boolean - If equals 1, then allow the Pierce modifier function to work with this missile. If equals 0, then do not allow Pierce to work with this missile.";
            missiles["MissileSkill"] = "Boolean - If equals 1, then the missile will look up the skill that created it and use that skill's damage instead of the missile damage. If equals 0, then ignore this.";
            missiles["Skill"] = "Open - Links to the \"skill\" field from the Skills.txt file. This will look up the specified skill's damage and use it for the missile instead of using the missile's defined damage.";
            missiles["ResultFlags"] = "Number - Controls different flags that can affect how the target reacts after being hit by the missile. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 5 (binary = 101) then that returns true for both the 4 (binary = 100) and 1 (binary = 1) bit field values. Referenced by the Bit Field Value of the Result Flags Table.";
            missiles["HitFlags"] = "Number - Controls different flags that can affect the damage dealt when the target is hit by the missile. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 6 (binary = 110) then that returns true for both the 4 (binary = 100) and 2 (binary = 10) bit field values. Referenced by the Bit Field Value of the Hit Flags Table.";
            missiles["HitShift"] = "Number - Controls the percentage modifier for the missile's damage. This value cannot be less than 0 or greater than 8. This is calculated in 256ths, where 8=256/256, 7=128/256, 6=64/256, 5=32/256, 4=16/256, 3=8/256, 2=4/256, 1=2/256, and 0=1/256.";
            missiles["ApplyMastery"] = "Boolean - If equals 1, then apply the caster's elemental mastery bonus modifiers to the missile's elemental damage. If equals 0, then ignore this.";
            missiles["SrcDamage"] = "Number - Controls how much of the source unit's damage should be added to the missile's damage. This is calculated in 128ths and acts as a percentage modifier for the source unit's damage that added to the missile. If equals -1 or 0, then the source damage is not included.";
            missiles["Half2HSrc"] = "Boolean - If equals 1 and the source unit is currently wielding a 2-Handed weapon, then the SrcDamage is reduced by 50%. If equals 0, then ignore this.";
            missiles["SrcMissDmg"] = "Number - If the missile was created by another missile, then this controls how much of the source missile's damage should be added to this missile's damage. This is calculated in 128ths and acts as a percentage modifier for the source missile's damage that added to this missile. If equals 0, then the source damage is not included.";
            missiles["MinDamage"] = "Number - Minimum baseline physical damage dealt by the missile.";
            missiles["MinLevDam1"] = "Number - Controls the additional minimum physical damage dealt by the missile, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field.";
            missiles["MaxDamage"] = "Number - Maximum baseline physical damage dealt by the missile.";
            missiles["MaxLevDam1"] = "Number - Controls the additional maximum physical damage dealt by the missile, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field.";
            missiles["DmgSymPerCalc"] = "Calculation - Determines the percentage increase to the physical damage dealt by the missile based on specified skill levels.";
            missiles["EType"] = "Open - Defines the type of elemental damage dealt by the missile. If this field is empty, then the related elemental fields below will not be used. Referenced by the Code value of the Elemental Types Table.";
            missiles["EMin"] = "Number - Minimum baseline elemental damage dealt by the missile.";
            missiles["MinELev1"] = "Number - Controls the additional minimum elemental damage dealt by the missile, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field number.";
            missiles["EMax"] = "Number - Maximum baseline elemental damage dealt by the missile.";
            missiles["MaxELev1"] = "Number - Controls the additional maximum elemental damage dealt by the missile, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field.";
            missiles["EDmgSymPerCalc"] = "Calculation - Determines the percentage increase to the elemental damage dealt by the missile based on specified skill levels.";
            missiles["ELen"] = "Number - The baseline elemental duration dealt by the missile. This is calculated in frame lengths where 25 Frames = 1 second. These fields only apply to appropriate elemental types with a duration.";
            missiles["ELevLen1"] = "Number - Controls the additional elemental duration added by the missile, calculated using the leveling formula between 3 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17 and beyond. These 3 level thresholds correlate to each field. These fields only apply to appropriate elemental types with a duration.";
            missiles["HitClass"] = "Number - Defines the missile's own hit class into the damage routines, mainly used for determining hit sound effects and overlays. This field only handles the hit class layers, so values beyond these defined bits are ignored. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 6 (binary = 110) then that returns true for both the 4 (binary = 100) and 2 (binary = 10) bit field values.";
            missiles["NumDirections"] = "Number - The number of directions allowed by the missile, based on the DCC file used (CelFile). This value should be within the power of 2, with a minimum value of 1 or up to a maximum value of 64.";
            missiles["LocalBlood"] = "Boolean - If equals 1, then change the color of blood missiles to green. If equals 0, then keep the blood missiles colored the default red.";
            missiles["DamageRate"] = "Number - Controls the \"damage_framerate\" stat (Calculated in 1024ths), which acts as a percentage multiplier for the physical damage reduction and magic damage reduction stat modifiers, when performing damage resistance calculations. This is only enabled if the value is greater than 0.";
            missiles["TravelSound"] = "Open - Points to a Sound field defined in the Sounds.txt file. Used when the missile is created and while it is alive.";
            missiles["HitSound"] = "Open - Points to a Sound field defined in the Sounds.txt file. Used when the collides with a target.";
            missiles["ProgSound"] = "Open - Points to a Sound field defined in the Sounds.txt file. Used for a programmed special event based on the client function.";
            missiles["ProgOverlay"] = "Open - Points to the overlay field defined in the Overlay.txt file. Used for a programmed special event based on the server or client function.";
            missiles["ExplosionMissile"] = "Open - Points to the Missile field for another missile. Used for the missile created on the client when this missile explodes.";
            missiles["SubMissile1"] = "Open - Points to the Missile field for another missile. Used for creating a new missile based on the server function used.";
            missiles["HitSubMissile1"] = "Open - Points to the Missile field for another missile. Used for a new missile after a collision, based on the server function used.";
            missiles["CltSubMissile1"] = "Open - Points to the Missile field for another missile. Used for creating a new missile based on the client function used.";
            missiles["CltHitSubMissile1"] = "Open - Points to the Missile field for another missile. Used for a new missile after a collision, based on the client function used.";

            monequip["monster"] = "Open - Defines the monster that should be equipped. Points to the matching ID value in the MonStats.txt file. If the monster has multiple defined equipment possibilities, then they should always be grouped together. The game will go through the list in order to match what is best to use for the monster.";
            monequip["oninit"] = "Boolean - Defines if the monster equipment is added on initialization during the monster's creation, depending how the monster is spawned. Monsters created by a skill have this value set to 0. Monsters created by a level have this value set to 1.";
            monequip["level"] = "Number - Defines the level requirement for the monster in order to gain this equipment. The game will prefer the highest level allowed, so the order of these equipment should be from highest level to lowest level.";
            monequip["item1"] = "Open - Item code that can be equipped on the monster.";
            monequip["loc1"] = "Open - Specifies the inventory slot where the item will be equipped. Once an item is equipped on that body location, then the game will skip any duplicate calls to equipping the same body location. This is another reason why the equipment should be ordered from highest level to lowest level. Referenced from the Code column in BodyLocs.txt.";
            monequip["mod1"] = "Number - Controls the quality level of the related item.";

            monlvl["Level"] = "Number - An integer value to determine how to scale the monster's statistics when at a specific level.";
            monlvl["AC & AC(N) & AC(H)"] = "Number - Percentage multiplier for increasing the Monster's Defense (multiplies with the AC field)";
            monlvl["TH & TH(N) & TH(H)"] = "Number - Percentage multiplier for increasing the Monster's Attack Rating (multiplies with the A1TH and A2TH fields)";
            monlvl["HP & HP(N) & HP(H)"] = "Number - Percentage multiplier for increasing the Monster's Life (multiplies with the minHP and maxHP fields)";
            monlvl["DM & DM(N) & DM(H)"] = "Number - Percentage multiplier for increasing the Monster's Damage (multiplies with the A1MinD, A1MaxD, A2MinD, A2MaxD, El1MinD and El1MaxD fields)";
            monlvl["XP & XP(N) & XP(H)"] = "Number - Percentage multiplier for increasing the Experience provided to the player when killing the Monster (multiplies with the Exp fields)";
            monlvl["L-AC & L-AC(N) & L-AC(H)"] = "Number - Percentage multiplier for increasing the Monster's Defense (multiplies with the AC field)";
            monlvl["L-TH & L-TH(N) & L-TH(H)"] = "Number - Percentage multiplier for increasing the Monster's Attack Rating (multiplies with the A1TH and A2TH fields)";
            monlvl["L-HP & L-HP(N) & L-HP(H)"] = "Number - Percentage multiplier for increasing the Monster's Life (multiplies with the minHP and maxHP fields)";
            monlvl["L-DM & L-DM(N) & L-DM(H)"] = "Number - Percentage multiplier for increasing the Monster's Damage (multiplies with the A1MinD, A1MaxD, A2MinD, A2MaxD, El1MinD and El1MaxD fields)";
            monlvl["L-XP & L-XP(N) & L-XP(H)"] = "Number - Percentage multiplier for increasing the Experience provided to the player when killing the Monster (multiplies with the Exp fields)";

            monpreset["Act"] = "Number - Defines the Act number used for each Monster Preset. Uses values between 1 to 5.";
            monpreset["Place"] = "Open - Defines a Super Unique monster from SuperUniques.txt, a monster from MonStats.txt or a place from MonPlace.txt. This defines the Monster Preset which is used for preloading, such as during level transitions.";

            monprop["ID"] = "Number - Defines the monster that should gain the Property. Points to the matching \"ID\" value in the monstats.txt file.";
            monprop["prop1"] = "Open - Defines with Property to apply to the monster (Uses the \"code field from Properties.txt).";
            monprop["chance1"] = "Number - The percent chance that the related property (prop#) will be assigned. If this value equals 0, then the Property will always be applied.";
            monprop["par1"] = "Number - The \"parameter\" value associated with the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";
            monprop["min1"] = "Number - The \"min\" value to assign to the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";
            monprop["max1"] = "Number - The \"max\" value to assign to the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";

            monseq["sequence"] = "Number - Establishes the Monster Sequence index. An entire monster sequence can be composed of multiple sequence lines, which means that each line needs to have matching \"sequence\" fields and must be in contiguous order.";
            monseq["mode"] = "Open - Defines which monster mode animation to use for the sequence. Referenced from the Code column in MonMode.txt.";
            monseq["frame"] = "Number - The in-game frame number for the animation. For the first line in the sequence, this value will establish where the starting frame for the animation. These values should be in contiguous order for the sequence.";
            monseq["dir"] = "Number - Defines the numeric animation direction that the frame use. Most animations have between 8 to 64 maximum directions.";
            monseq["event"] = "Number - Defines what type of event will be used when the frame triggers.";

            monstats["ID"] = "Number - Controls the unique name ID to define the monster.";
            monstats["BaseID"] = "Open - Points to the ID of another monster to define the monster's base type. This is to create groups of monsters which are considered the same type.";
            monstats["NextInClass"] = "Open - Points to the ID of another monster to signify the next monster in the group of this monster's type. This is to continue the groups of monsters which are considered the same type. The order should be contiguous.";
            monstats["TransLvl"] = "Number - Defines the color transform level to use for this monster, which affects what color palette that the monster will use.";
            monstats["MonStatsEx"] = "Open - Controls a pointer to the ID from MonStats2.txt.";
            monstats["MonProp"] = "Open - Points to the ID field from MonProp.txt. Used to add special modifiers to the monster.";
            monstats["MonType"] = "Open - Points to the type field from MonType.txt. Used to handle the monster's classification.";
            monstats["AI"] = "Open - Points to a type of AI script to use for the monster (MonAI.txt).";
            monstats["Code"] = "Open - Controls the token used for choosing the proper cells to display the monster's graphics.";
            monstats["enabled"] = "Boolean - If equals 1, then this monster is allowed to spawn in the game. If equals 0, then this monster will never spawn in the game.";
            monstats["rangedtype"] = "Boolean - If equals 1, then the monster will be classified as a ranged type. If equals 0, then the monster will be classified as a melee type.";
            monstats["placespawn"] = "Boolean - If equals 1, then this monster will be treated as a spawner, so monsters that spawn can be initially placed within this monster. If equals 0, then ignore this.";
            monstats["spawn"] = "Open - Points to the ID of another monster to control what kind of monster is spawned from this monster. This is only used if the placespawn field is enabled.";
            monstats["spawnmode"] = "Open - Defines the animation mode that the spawned monsters will be initiated with. Referenced from the Token column in MonMode.txt.";
            monstats["minion1"] = "Open - Points to the ID of another monster to control what kind of monster is spawned with this monster when it is spawned, like a monster pack. The minion1 field is also used for spawning a monster when this monster is killed while it has the SplEndDeath field enabled.";
            monstats["SetBoss"] = "Boolean - If equals 1, then set the monster AI to use the Boss AI type, which can affect the monster's behaviors. If equals 0, then ignore this.";
            monstats["BossXfer"] = "Boolean - If equals 1, then the monster's AI will transfer its boss recognition to another monster, which can affect the minion monster behaviors after this boss is killed. If equals 0, then ignore no boss AI will transfer and minion monsters will behave differently after the boss is killed. This field relies on the SetBoss field being enabled.";
            monstats["PartyMin"] = "Number - The minimum number of minions that can spawn with this monster. Uses the minion1 fields. The actual number is a random value chosen between the \"PartyMin\" and \"PartyMax\" field values.";
            monstats["PartyMax"] = "Number - The maximum number of minions that can spawn with this monster. Uses the minion1 fields. The actual number is a random value chosen between the \"PartyMin\" and \"PartyMax\" field values.";
            monstats["MinGrp"] = "Number - The minimum number of duplicates of this monster that can spawn together. The actual number is a random value chosen between the \"MinGrp\" and \"MaxGrp\" field values.";
            monstats["MaxGrp"] = "Number - The maximum number of duplicates of this monster that can spawn together. The actual number is a random value chosen between the \"MinGrp\" and \"MaxGrp\" field values.";
            monstats["sparsePopulate"] = "Number - If this value is greater than 0, then it controls the percent chance that this monster does not spawn, and another monster will spawn in its place. (Out of 100).";
            monstats["Velocity"] = "Number - Determines the movement velocity of the monster, which can be the monster's baseline walk speed.";
            monstats["Run"] = "Number - Determines the run speed of the monster as opposed to walk speed. This is only used if the monster has a Run mode.";
            monstats["Rarity"] = "Number - Modifies the chance that this monster will be chosen to spawn in the area level. The higher the value is, then the more likely this monster will be chosen. This value acts as a numerator and a denominator. All \"Rarity\" values of possible monsters get summed together to give a total denominator, used for the random roll. For example, if there are 3 possible monsters that can spawn, and their \"Rarity\" values are 1, 2, 2, then their chances to be chosen are 1/5, 2/5, and 2/5 respectively. If this value equals 0, then this monster is never randomly selected to spawn in an area level.";
            monstats["Level"] = "Number - Determines the monster's level. This value for Nightmare and Hell difficulty can be overridden by the area level's MonLvl or MonLvlEx values from Levels.txt, unless the monster's boss and noRatio fields are enabled.";
            monstats["MonSound"] = "Open - Points to the ID field of a monster sound from MonSounds.txt. This is used to control the monsters assigned sounds, when the monster is spawned as a Normal monster.";
            monstats["UMonSound"] = "Open - Points to the ID field of a monster sound from MonSounds.txt. This is used to control the monsters assigned sounds, when the monster is spawned as a Unique or Champion monster.";
            monstats["threat"] = "Number - Controls the AI threat value of the monster which can affect the targeting priorities of enemy Ais for this monster. The higher this value is, then the more likely that enemy AI will target this monster.";
            monstats["aidel"] = "Number - Controls the delay in frame length for how often the monster's AI will update its commands. A lower delay means that the monster will perform commands more often without as long of a pause in between.";
            monstats["aidist"] = "Number - Controls the maximum distance (measured in tiles) between the monster and an enemy until the monster's AI becomes aggressive. If equals 0, then default to 35.";
            monstats["aip1"] = "Number - Defines numeric parameters used to control various functions of the monster's AI. These fields depend on which AI script is being used (MonAI.txt, and the AI field in MonStats.txt).";
            monstats["MissA1"] = "Open - Points to the Missile field from Missiles.txt to determine which missile to use when the monster is in Attack 1 & Attack 2 mode.";
            monstats["MissS1"] = "Open - Points to the Missile field from Missiles.txt to determine which missile to use when the monster is in Skill 1 (to Skill 4) mode.";
            monstats["MissC"] = "Open - Points to the Missile field from Missiles.txt to determine which missile to use when the monster is in Cast mode.";
            monstats["MissSQ"] = "Open - Points to the Missile field from Missiles.txt to determine which missile to use when the monster is in Sequence mode.";
            monstats["Align"] = "Number - Controls the monster's alignment, which determines if the monster will be an enemy, ally, or neutral party to the player.";
            monstats["isSpawn"] = "Boolean - If equals 1, then the monster is allowed to spawn in an area level. If equals 0, then the monster will not be spawned automatically in an area level.";
            monstats["isMelee"] = "Boolean - If equals 1, then the monster is classified as a melee only type, which can affect its AI behaviors and what monster modifiers are allowed on the monster. If equals 0, then ignore this.";
            monstats["npc"] = "Boolean - If equals 1, then the monster is classified as an NPC (Non-Playable Character), which can affect its AI behaviors and how the player treats this monster. If equals 0, then ignore this.";
            monstats["interact"] = "Boolean - If equals 1, then the monster is interactable, meaning that the player can click on the monster to perform an interact command instead of attacking. If equals 0, then ignore this.";
            monstats["inventory"] = "Boolean - If equals 1, then monster will have an inventory with randomly generated items, such as an NPC with shop items (if the interact field is enabled) or a summoned unit with random equipped items (MonEquip.txt). If equals 0, then ignore this.";
            monstats["inTown"] = "Boolean - If equals 1, then the monster is allowed to be in town. If equals 0, then the monster is not allowed to be in town, which can affect or disable their AI or collision from entering towns.";
            monstats["lUndead"] = "Boolean - If equals 1, then the monster is treated as a Low Undead, meaning that the monster is classified as an Undead type and can be resurrected by certain AI. If equals 0, then ignore this.";
            monstats["hUndead"] = "Boolean - If equals 1, then the monster is treated as a High Undead, meaning that the monster is classified as an Undead type but cannot be resurrected by certain AI. If equals 0, then ignore this.";
            monstats["demon"] = "Boolean - If equals 1, then the monster is classified as a Demon type. If equals 0, then ignore this.";
            monstats["flying"] = "Boolean - If equals 1, then the monster is flagged as a flying type, which can affect its collision with the area level and how it is spawned. If equals 0, then ignore this.";
            monstats["opendoors"] = "Boolean - If equals 1, then the monster will use its AI to open doors if necessary. If equals 0, then the monster cannot open doors and will treat doors as another type of collision.";
            monstats["boss"] = "Boolean - If equals 1, then the monster is classified as a Boss type, which can affect boss related AI and functions. If equals 0, then ignore this.";
            monstats["primeevil"] = "Boolean - If equals 1, then the monster is classified as a Prime Evil type, or an Act End boss, which can affect various skills, AI, and damage related functions. If equals 0, then ignore this.";
            monstats["killable"] = "Boolean - If equals 1, then the monster can be killed, damage, and be put in a Death or Dead mode. If equals 0, then the monster cannot be damaged or killed.";
            monstats["switchai"] = "Boolean - If equals 1, then monster's AI can switched, such as by the Assassin's Mind Blast ability. If equals 0, then the monster AI cannot be switched.";
            monstats["noAura"] = "Boolean - If equals 1, then the monster cannot be affected by friendly auras. If equals 0, then the monster can be affected by friendly auras.";
            monstats["nomultishot"] = "Boolean - If equals 1, then the monster is not allowed to spawn with the Multi-Shot unique monster modifier (MonUMod.txt). If equals 0, then ignore this.";
            monstats["neverCount"] = "Boolean - If equals 1, then the monster is not counted on the list of the active monsters in the area, which affects spawning and saving functions. If equals 0, then the monster will be accounted for, and can be part of the active or inactive list functions.";
            monstats["petIgnore"] = "Boolean - If equals 1, then pet AI scripts will ignore this monster (MonAI.txt). If equals 0, then pet AI will attack this monster.";
            monstats["deathDmg"] = "Boolean - If equals 1, then the monster will explode on death. This has special cases for the \"bonefetish1\" and \"siegebeast1\" monster classes, otherwise the monster will use a general death damage function to damage nearby units based on the monster's health percentage. If equals 0, then ignore this.";
            monstats["genericSpawn"] = "Boolean - If equals 1, the monster is flagged as a possible selection for the AI generic spawner function. There are defaults for using the If equals 0, then ignore this.";
            monstats["zoo"] = "Boolean - If equals 1, then the monster will be flagged as a zoo type monster, which will give it the AI zoo behavior. If equals 0, then ignore this.";
            monstats["CannotDesecrate"] = "Boolean - If equals 1, then the monster will not be able to be desecrated when inside a desecrated level. If equals 0, then ignore this.";
            monstats["rightArmItemType"] = "Open - Determines what type of items the monster is allowed to hold in its right arm (ItemTypes.txt). A blank value means it can hold any item.";
            monstats["leftArmItemType"] = "Open - Determines what type of items the monster is allowed to hold in its left arm (ItemTypes.txt). A blank value means it can hold any item.";
            monstats["canNotUseTwoHandedItems"] = "Boolean - If equals 1, then the monster can not items marked as two handed (Weapons.txt).";
            monstats["SendSkills"] = "Number - Determines which of the monster's skill's level should be sent to the client. Uses a byte value, where the code tests each bit to determine which of the monster's skills to check.";
            monstats["Skill1"] = "Open - Points to the Skill field from Skills.txt file. This gives the monster the skill to use for Sk1Mode.";
            monstats["Sk1mode"] = "Open - Determines the monster's animation mode when using the related skill. Outside of the standard animation mode inputs, the field can also point to a sequence from MonSeq.txt, which handles a specific set of frames to place a sequence animation. Referenced from the Code column in MonMode.txt.";
            monstats["Sk1lvl"] = "Number - Controls the base skill level of the related skill on the monster.";
            monstats["Drain"] = "Number - Controls the monster's overall Life and Mana steal percentage. This can also be affected by the LifeStealDivisor and ManaStealDivisor fields from DifficultyLevels.txt. If equals 0, then the monster will not have Life or Mana steal.";
            monstats["coldeffect"] = "Number - Sets the percentage change in movement speed and attack rate when the monster if chilled by a cold effect. If this equals 0, then the monster will be immune to the cold effect.";
            monstats["ResDm"] = "Number - Sets the monster's Physical Damage Resistance stat.";
            monstats["ResMa"] = "Number - Sets the monster's Magic Resistance stat.";
            monstats["ResFi"] = "Number - Sets the monster's Fire Resistance stat.";
            monstats["ResLi"] = "Number - Sets the monster's Lightning Resistance stat.";
            monstats["ResCo"] = "Number - Sets the monster's Cold Resistance stat.";
            monstats["ResPo"] = "Number - Sets the monster's Poison Resistance stat.";
            monstats["DamageRegen"] = "Number - Controls the monster's Life regeneration per frame. This is calculated based on the monster's maximum life: Regeneration Rate = (Life * \"DamageRegen\") / 16.";
            monstats["SkillDamage"] = "Number - Points to a skill from the \"skill\" field in the Skills.txt file. This changes the monster's min physical damage, max physical damage, and Attack Rating to be based off the values from the linked skill and its current level from the monster's owner (usually the player who summoned the monster).";
            monstats["noRatio"] = "Boolean - If equals 1, then use this file's fields to determine the monster's baseline stats (minHP, maxHP, AC, Exp, A1MinD, A1MaxD, A1TH, A2MinD, A2MinD, A2TH, S1MinD, S1MaxD, S1TH). If equals 0, then use MonLvl.txt to determine the monster's baseline stats.";
            monstats["ShieldBlockOverride"] = "Number - If equals 1, then the monster can block without a shield (the block chance stat will take effect even without a shield equipped). If equals 2, then the monster cannot block at all, even with a shield equipped. If equals 0, then ignore this.";
            monstats["ToBlock"] = "Number - The monster's percent chance to block an attack.";
            monstats["Crit"] = "Number - The percent chance for the monster to score a critical hit when attacking an enemy, which causes the attack to deal double damage.";
            monstats["minHP"] = "Number - The monster's minimum amount of Life when spawned.";
            monstats["maxHP"] = "Number - The monster's maximum amount of Life when spawned.";
            monstats["AC"] = "Number - The monster's Defense value.";
            monstats["Exp"] = "Number - The amount of Experience that is rewarded to the player when the monster is killed.";
            monstats["A1MinD"] = "Number - The minimum damage dealt by the monster when it is using the Attack 1 (A1) animation mode.";
            monstats["A1MaxD"] = "Number - The maximum damage dealt by the monster when it is using the Attack 1 (A1) animation mode.";
            monstats["A1TH"] = "Number - The monster's Attack Rating when it is using the Attack 1 (A1) animation mode.";
            monstats["A2MinD"] = "Number - The minimum damage dealt by the monster when it is using the Attack 2 (A2) animation mode.";
            monstats["A2MaxD"] = "Number - The maximum damage dealt by the monster when it is using the Attack 2 (A2) animation mode.";
            monstats["A2TH"] = "Number - The monster's Attack Rating when it is using the Attack 2 (A2) animation mode.";
            monstats["S1MinD"] = "Number - The minimum damage dealt by the monster when it is using the Skill 1 (S1) animation mode.";
            monstats["S1MaxD"] = "Number - The maximum damage dealt by the monster when it is using the Skill 1 (S1) animation mode.";
            monstats["S1TH"] = "Number - The monster's Attack Rating when it is using the Skill 1 (S1) animation mode.";
            monstats["El1Mode"] = "Open - Determines which animation mode will trigger an additional elemental damage type when used. Referenced from the Code column in MonMode.txt.";
            monstats["El1Type"] = "Open - Defines the type of elemental damage. This field is used when El#Mode is not null. Referenced by the Code value of the Elemental Types Table.";
            monstats["El1Pct"] = "Number - Controls the random percent chance (out of 100) that the monster will append the element damage to the attack. This field is used when El#Mode is not null.";
            monstats["El1MinD"] = "Number - The minimum element damage applied to the attack. This field is used when El#Mode is not null.";
            monstats["El1MaxD"] = "Number - The maximum element damage applied to the attack. This field is used when El#Mode is not null.";
            monstats["El1Dur"] = "Number - Controls the duration of the related element mode in frame lengths (25 Frames = 1 Second). This is only applicable for the Cold, Poison, Stun, Burning, Freeze elements. There are special cases when evaluating the elements, where Poison min and max damage are multiplied by 10, and Poison duration is multiplied by 2. This field is used when El#Mode is not null.";
            monstats["TreasureClass"] = "Open - Defines which Treasure Class is used by the monster when a Normal monster type is killed; for each difficulty respectively.";
            monstats["TreasureClassChamp"] = "Open - Defines which Treasure Class is used by the monster when a Champion monster type is killed; for each difficulty respectively.";
            monstats["TreasureClassUnique"] = "Open - Defines which Treasure Class is used by the monster when a Unique monster type is killed; for each difficulty respectively.";
            monstats["TreasureClassQuest"] = "Open - Defines which Treasure Class is used by the monster when a Quest-Enabled monster type is killed; for each difficulty respectively. Based on the TCQuestID and TCQuestCP fields.";
            monstats["TCQuestID"] = "Number - Checks to see if the player has does not have a quest flag progress. If not, then use the TreasureClassQuest field, based on the game's current difficulty. Referenced by the Code value of the Quest Flags Table.";
            monstats["TCQuestCP"] = "Number - Controls which Quest Checkpoint, or current progress within a quest (based on the TCQuestID value), is needed to use the TreasureClassQuest field, based on the game's current difficulty.";
            monstats["SplEndDeath"] = "Number - Controls a special case death handler for the monster that is ran on the server side.";
            monstats["SplGetModeChart"] = "Boolean - If equals 1, then check special case handlers of certain monsters with specific BaseID fields while they are using certain a mode and perform a function. If equals 0, then ignore this.";
            monstats["SplEndGeneric"] = "Boolean - If equals 1, then check special case handlers of monsters with specific BaseID fields while they are ending certain modes and perform a function. If equals 0, then ignore this.";
            monstats["SplClientEnd"] = "Boolean - If equals 1, then on the client side, check special case handlers of monsters with specific BaseID fields while they are ending certain modes and perform a function. If equals 0, then ignore this.";

            monstats2["ID"] = "Open - Controls the unique name ID to define the monster. This must match the same value in the monstats.txt file.";
            monstats2["Height"] = "Number - Determines the height of the monster. This has 2 purposes. The first purpose is to act as an index value for selecting which icebreak missile to use when the monster dies while frozen. The second purpose is to select a code which affects what attack animation the player characters will use when attacking the monster (Attack1 and or Attack2). See each code description types below.";
            monstats2["OverlayHeight"] = "Number - Determines the height value of overlays (Overlay.txt) for the monster.";
            monstats2["pixHeight"] = "Number - Determines the pixel height value for the damage bar when the monster is selected.";
            monstats2["SizeX & SizeY"] = "Number - Determines the tile grid size of the monster which is use for handling placement when the monster spawns or uses movement skills";
            monstats2["spawnCol"] = "Number - Controls the method for spawning the monster based on the collisions in the environment.";
            monstats2["MeleeRng"] = "Number - Controls the range of the monster's melee attack, which can affect also affect certain AI pathing. If this value equals 255, then refer to the monster's weapon class (\"BaseW\").";
            monstats2["BaseW"] = "Open - Defines the monster's base weapon class, which can affect how the monster attacks. Referenced by the Code value of the Weapon Class Table.";
            monstats2["HitClass"] = "Open - Defines the specific class of an attack when the monster successfully hits with an attack. This can affect the sound and overlay display of the attack hit. Referenced from the Index value in HitClass.txt.";
            monstats2["HDv"] = "Open - Head visual.";
            monstats2["TRv"] = "Open - Torso visual.";
            monstats2["LGv"] = "Open - Legs visual.";
            monstats2["RAv"] = "Open - Right Arm visual.";
            monstats2["LAv"] = "Open - Left Arm visual.";
            monstats2["RHv"] = "Open - Right Hand visual.";
            monstats2["LHv"] = "Open - Left Hand visual.";
            monstats2["SHv"] = "Open - Shield visual.";
            monstats2["S1v (to S8v)"] = "Open - Special 1 to Special 8 visual.";
            monstats2["HD"] = "Boolean - Head.";
            monstats2["TR"] = "Boolean - Torso.";
            monstats2["LG"] = "Boolean - Legs.";
            monstats2["RA"] = "Boolean - Right Arm.";
            monstats2["LA"] = "Boolean - Left Arm.";
            monstats2["RH"] = "Boolean - Right Hand.";
            monstats2["LH"] = "Boolean - Left Hand.";
            monstats2["SH"] = "Boolean - Shield.";
            monstats2["S1"] = "Boolean - Special 1 to Special 8.";
            monstats2["TotalPieces"] = "Number - Defines the total amount of component pieces that the monster uses. This value should match the number of enabled Boolean fields listed above.";
            monstats2["mDT"] = "Boolean - If equals 1, then enable the Death Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mNU"] = "Boolean - If equals 1, then enable the Neutral Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mWL"] = "Boolean - If equals 1, then enable the Walk Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mGH"] = "Boolean - If equals 1, then enable the Get Hit Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mA1"] = "Boolean - If equals 1, then enable the Attack 1 (and Attack 2) Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mBL"] = "Boolean - If equals 1, then enable the Block Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mSC"] = "Boolean - If equals 1, then enable the Cast Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mS1"] = "Boolean - If equals 1, then enable the Skill 1 (to Skill4) Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mDD"] = "Boolean - If equals 1, then enable the Dead Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mKB"] = "Boolean - If equals 1, then enable the Knockback Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mSQ"] = "Boolean - If equals 1, then enable the Sequence Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["mRN"] = "Boolean - If equals 1, then enable the Run Mode for the monster. If equals 0, then this mode is disabled.";
            monstats2["dDT"] = "Number - Defines the number of directions that the monster can face during Death Mode.";
            monstats2["dNU"] = "Number - Defines the number of directions that the monster can face during Neutral Mode.";
            monstats2["dWL"] = "Number - Defines the number of directions that the monster can face during Walk Mode.";
            monstats2["dGH"] = "Number - Defines the number of directions that the monster can face during Get Hit Mode.";
            monstats2["dA1"] = "Number - Defines the number of directions that the monster can face during Attack 1 (and Attack 2) Mode.";
            monstats2["dBL"] = "Number - Defines the number of directions that the monster can face during Block Mode.";
            monstats2["dSC"] = "Number - Defines the number of directions that the monster can face during Cast Mode.";
            monstats2["dS1"] = "Number - Defines the number of directions that the monster can face during Skill 1 (to Skill 4) Mode.";
            monstats2["dDD"] = "Number - Defines the number of directions that the monster can face during Dead Mode.";
            monstats2["dKB"] = "Number - Defines the number of directions that the monster can face during Knockback Mode.";
            monstats2["dSQ"] = "Number - Defines the number of directions that the monster can face during Sequence Mode.";
            monstats2["dRN"] = "Number - Defines the number of directions that the monster can face during Run Mode.";
            monstats2["A1mv"] = "Boolean - If equals 1, then enable the Attack 1 (and Attack 2) Mode while the monster is moving with the Walk mode or Run mode. If equals 0, then this mode is disabled while the monster is moving.";
            monstats2["SCmv"] = "Boolean - If equals 1, then enable the Cast Mode while the monster is moving with the Walk mode or Run mode. If equals 0, then this mode is disabled while the monster is moving.";
            monstats2["S1mv"] = "Boolean - If equals 1, then enable the Skill 1 (to Skill 4) Mode while the monster is moving with the Walk mode or Run mode. If equals 0, then this mode is disabled while the monster is moving.";
            monstats2["noGfxHitTest"] = "Boolean - If equals 1, then enable the mouse selection bounding box functionality around the monster. If equals 0, then the monster cannot be selected by the mouse.";
            monstats2["htTop"] = "Number - Define the pixel top offset around the monster for the mouse selection bounding box functionality. This field relies on the noGfxHitTest field being enabled.";
            monstats2["htLeft"] = "Number - Define the pixel left offset around the monster for the mouse selection bounding box functionality. This field relies on the noGfxHitTest field being enabled.";
            monstats2["htWidth"] = "Number - Define the pixel right offset around the monster for the mouse selection bounding box functionality. This field relies on the noGfxHitTest field being enabled.";
            monstats2["htHeight"] = "Number - Define the pixel bottom offset around the monster for the mouse selection bounding box functionality. This field relies on the noGfxHitTest field being enabled.";
            monstats2["restore"] = "Number - Determines if the monster should be placed on the inactive list, to be saved when the level unloads. If equals 0, then do not save the monster. If equals 1, then rely on other checks to determine to save the monster. If equals 2, then force save the monster.";
            monstats2["automapCel"] = "Number - Controls what index of the Automap tiles to use to display this monster on the Automap. This field relies on the noMap field being disabled.";
            monstats2["noMap"] = "Boolean - If equals 1, then the monster will not appear on the Automap. If equals 0, then the monster will normally appear on the Automap.";
            monstats2["noOvly"] = "Boolean - If equals 1, then no looping overlays will be drawn on the monster. If equals 0, then overlays will be drawn on the monster. (Overlay.txt).";
            monstats2["isSel"] = "Boolean - If equals 1, then the monster is selectable and can be targeted. If equals 0, then the monster cannot be selected.";
            monstats2["alSel"] = "Boolean - If equals 1, then the player can always select the monster, regardless of being an ally or enemy. If equals 0, then ignore this.";
            monstats2["noSel"] = "Boolean - If equals 1, then the player can never select the monster. If equals 0, then ignore this.";
            monstats2["shiftSel"] = "Boolean - If equals 1, then the player can target this monster when holding the Shift key and clicking to use a skill. If equals 0, then the monster cannot be targeted while the player is holding the Shift key.";
            monstats2["corpseSel"] = "Boolean - If equals 1, then the monster's corpse can be with the mouse cursor. If equals 0, then the monster's corpse cannot be selected with the mouse cursor.";
            monstats2["isAtt"] = "Boolean - If equals 1, then the monster can be attacked. If equals 0, then the monster cannot be attacked.";
            monstats2["revive"] = "Boolean - If equals 1, then the monster is allowed to be revived by the Necromancer Revive skill. If equals 0, then the monster cannot be revived by the Necromancer Revive skill.";
            monstats2["limitCorpses"] = "Boolean - If equals 1, then the monster's corpse will be placed into a pool with all other corpses with this field checked. Once that pool reaches the max (50), the corpses at the beginning of the pool start getting removed.";
            monstats2["critter"] = "Boolean - If equals 1, then the monster will be flagged as a critter, which gives some special case handling such as not creating impact sounds and differently handling its spawn placement in presets. If equals 0, then ignore this.";
            monstats2["small"] = "Boolean - If equals 1, then the monster will be classified as a small type, which can affect what types of missiles can be used on the monster (Example: Barbarian Grim Ward size) or how the monster is knocked back. If equals 0, then ignore this. If this field is enabled, then the large field should be disabled, to avoid confusion.";
            monstats2["large"] = "Boolean - If equals 1, then the monster will be classified as a large type, which can affect what types of missiles can be used on the monster (Example: Barbarian Grim Ward size) or how the monster is knocked back. If equals 0, then ignore this. If this field is enabled, then the small field should be disabled, to avoid confusion.";
            monstats2["soft"] = "Boolean - If equals 1, then the monster's corpse is classified as soft bodied, meaning that its corpse can be used by certain corpse skills such as Barbarian Find Potion, Find Item, or Grim Ward. If equals 0, then the monster's corpse cannot be used for these skills.";
            monstats2["inert"] = "Boolean - If equals 1, then the monster will never attack its enemies. If equals 0, then ignore this.";
            monstats2["objCol"] = "Boolean - If equals 1 and the monster class is \"barricadedoor\", \"barricadedoor2\", or \"evilhut\", then the monster will place an invisible object with collision. If equals 0, then ignore this.";
            monstats2["deadCol"] = "Boolean - If equals 1, then the monster's corpse will have collision with other units. If equals 0, then the monster's corpse will not have collision.";
            monstats2["unflatDead"] = "Boolean - If equals 1, then ignore the corpse draw order for rendering the Sprite on top of others, while the monster is dead. If equals 0, then the monster's corpse will have a normal corpse draw order.";
            monstats2["Shadow"] = "Boolean - If equals 1, then the monster will project a shadow on the ground. If equals 0, then the monster will not project a shadow.";
            monstats2["noUniqueShift"] = "Boolean - If equals 1 and the monster is a Unique monster, then the monster will not have random color palette transform shifts. If equals 0, then the non-Unique monster will have random color palette transform shifts.";
            monstats2["compositeDeath"] = "Boolean - If equals 1, then the monster's Death Mode and Dead mode will make use of its component system. If equals 0, then the monster will default to using the Hand-To-Hand weapon class and no component system.";
            monstats2["localBlood"] = "Number - Controls the color of the monster's blood based on the region locale. If equals 0, then do not change the blood to green and keep it red. If equals 1, then change the monster's special components to use the green blood locale. If equals 2, then change the blood to green.";
            monstats2["Bleed"] = "Number - Controls if the monster will create blood missiles. If equals 0, then the monster will never bleed. If equals 1, then the monster will randomly create the \"blood1\" or \"blood2\" missiles when hit. If equals 2, then the monster will randomly create the \"blood1\", \"blood2\", \"bigblood1\", or \"bigblood2\" missiles when hit.";
            monstats2["Light"] = "Number - Controls the monster's minimum Light Radius size (measured in grid sub-tiles).";
            monstats2["light-r"] = "Number - Controls the red color value of the monster's Light Radius (Uses a value from 0 to 255)";
            monstats2["light-g"] = "Number - Controls the green color value of the monster's Light Radius (Uses a value from 0 to 255)";
            monstats2["light-b"] = "Number - Controls the blue color value of the monster's Light Radius (Uses a value from 0 to 255)";
            monstats2["Utrans & Utrans(N) & Utrans(H)"] = "Number - Modifies the color palette transform for the monster respectively in Normal, Nightmare, and Hell difficulty.";
            monstats2["InfernoLen"] = "Number - The frame length to hold the channel cast time of the inferno skill. This is used for when the monster has the \"inferno\" state, or for Diablo when he is using the \"DiabLight\" skill.";
            monstats2["InfernoAnim"] = "Number - The exact frame in the channel animation to loop back and start at again.";
            monstats2["InfernoRollback"] = "Number - The exact frame in the channel animation to determine when to roll back to the InfernoAnim frame.";
            monstats2["ResurrectMode"] = "Open - Controls which monster mode to set on the monster when it is resurrected. Referenced from the Code column in MonMode.txt.";
            monstats2["ResurrectSkill"] = "Open - Controls what skill should the monster use when it is resurrected (Skills.txt).";
            monstats2["SpawnUniqueMod"] = "Open - Controls what unique modifier the monster should always spawn with (MonUMod.txt).";

            montype["type"] = "Open - Defines the unique monster type ID.";
            montype["equiv1"] = "Open - Points to the index of another Monster Type to reference as a parent. This is used to create a hierarchy for Monster Types where the parents will have more universal settings shared across the related children.";
            montype["strplur"] = "Open - Uses a string for the plural form of the monster type. This is used by Function 22 of the descfunc field from ItemStatCost.txt, based on the monster type selected.";
            montype["element"] = "Open - Defines the monster's element type. This can be used for the Necromancer's Raise Skeletal Mage skill for determining what elemental type a Skeletal Mage should be based on the monster it was raised from (If the monster has no element, then the skeletal mage element will be randomly selected).";

            monumod["uniquemod"] = "Open - This is a reference field to define the monster modifier.";
            monumod["id"] = "Number - Defines the unique numeric ID for the monster modifier. Used as a reference in other data files.";
            monumod["enabled"] = "Boolean - If equals 1, then this monster modifier will be an available option for monsters to spawn with. If equals 0, then this monster modifier will never be used.";
            monumod["version"] = "Number - Defines which game version to use this monster modifier (<100 = Classic mode | 100 = Expansion mode).";
            monumod["xfer"] = "Boolean - If equals 1, then this monster modifier can be transferred from the Boss monster to his Minion monsters, including auras. If equals 0, then the monster modifier will never be transferred.";
            monumod["champion"] = "Boolean - If equals 1, then this monster modifier will only be used by Champion monsters. If equals 0, then the monster modifier can be used by any type of special monster.";
            monumod["fPick"] = "Number - Controls if this monster modifier is allowed on the monster based on the function code and the parameters it checks.";
            monumod["exclude1"] = "Number - This controls which Monster Types should not have this monster modifier (Uses the type field from MonType.txt).";
            monumod["cpick & cpick (N) & cpick (H)"] = "Number - Modifies the chances that this monster modifier will be chosen for a Champion monster, compared to other monster modifiers. The higher the value is, then the more likely this modifier will be chosen. This value acts as a numerator and a denominator. All 'cpick' values get summed together to give a total denominator, used for the random roll. For example, if there are 3 possible monster modifiers, and their 'cpick' values are 3, 4, 6, then their chances to be chosen are 3/13, 4/13, and 6/13 respectively";
            monumod["upick & upick (N) & upick (H)"] = "Number - Modifies the chances that this monster modifier will be chosen for a Unique monster, compared to other monster modifiers. The higher the value is, then the more likely this modifier will be chosen. This value acts as a numerator and a denominator. All 'upick' values get summed together to give a total denominator, used for the random roll. For example, if there are 3 possible monster modifiers, and their 'upick' values are 3, 4, 6, then their chances to be chosen are 3/13, 4/13, and 6/13 respectively";
            monumod["constants"] = "Number - These values control a special list of numeric parameters for special monsters. The row that each constant appears in the data file is unrelated. You can treat this column almost like a separate data file that controls other aspects of special monsters. See the description next to each value for more specific clarification on each constant.";

            monsounds["ID"] = "Open - Defines the unique name ID for the monster sound.";
            monsounds["Attack1"] = "Open - Play this sound when the monster performs Attack 1 and Attack 2, respectively.";
            monsounds["Weapon1"] = "Open - Play this sound when the monster performs Attack 1 and Attack 2, respectively. This acts as an extra sound that can play with the Attack1 sounds.";
            monsounds["Att1Del"] = "Number - Controls the amount of game frames to delay playing the Attack1 sounds, respectively.";
            monsounds["Wea1Del"] = "Number - Controls the amount of game frames to delay playing the Weapon1 sounds, respectively.";
            monsounds["Att1Prb"] = "Number - Controls the percent chance (out of 100) to play the Attack1 sounds, respectively.";
            monsounds["Wea1Vol"] = "Number - Controls the volume of the Weapon1 sounds, respectively. Uses a range between 0 to 255, where 255 is the maximum volume.";
            monsounds["HitSound"] = "Open - Play this sound when the monster gets hit or knocked back.";
            monsounds["DeathSound"] = "Open - Play this sound when the monster dies.";
            monsounds["HitDelay"] = "Number - Controls the amount of game frames to delay playing the HitSound.";
            monsounds["DeaDelay"] = "Number - Controls the amount of game frames to delay playing the DeathSound.";
            monsounds["Skill1"] = "Open - Play this sound when the monster uses the skill linked in the related Skill1 field from MonStats.txt.";
            monsounds["Footstep"] = "Open - Play this sound while the monster is walking or running.";
            monsounds["FootstepLayer"] = "Open - Play this sound while the monster is walking or running. This acts as an extra sound that can play with the Footstep sound.";
            monsounds["FsCnt"] = "Number - Controls the footstep count which is used to determine how often to play the Footstep and FootstepLayer sound. A higher value would mean that the sounds would play more often.";
            monsounds["FsOff"] = "Number - Controls the footstep offset which is used for calculating when to play the next Footstep and FootstepLayer sound, based on the current animation frame and the animation rate. A higher value would mean that the sounds would play less often.";
            monsounds["FsPrb"] = "Number - Controls the probability to play the Footstep and FootstepLayer sound, with a random chance out of 100.";
            monsounds["Neutral"] = "Open - Play this sound while the monster is in Neutral, Walk, or Run mode. Also play this sound when the monster ID equals \"vulture1\" while using Skill1 or the ID equals \"batdemon1\" while using Skill4.";
            monsounds["NeuTime"] = "Number - Controls the amount of game frames to delay between re-playing the Neutral sound after it finishes.";
            monsounds["Init"] = "Open - Play this sound when the monster spawns and is not dead and is not playing its Neutral sound. Sound entry from Sounds.txt.";
            monsounds["Taunt"] = "Open - Play this sound when the server requests that the monster should play its Taunt. This is typically used for quest or story related moments. Sound entry from Sounds.txt.";
            monsounds["Flee"] = "Open - Play this sound when the monster is told to flee. This depends on when the monster AI is told to play this sound. Sound entry from Sounds.txt.";
            monsounds["CvtMo1"] = "Open - This is used to convert the mode for playing the sound. This field defines the original mode that the monster is using. (MonMode.txt for the list of possible inputs).";
            monsounds["CvtSk1"] = "Open - Defines the skill that the monster is using. If the monster uses a specific skill, then the game can change the monster's mode for sound functionalities to another mode to change how sounds are generally handled. Points to a Skill in the Skills.txt file.";
            monsounds["CvtTgt1"] = "Open - Defines the mode to convert the sound to when the monster is using the relative skill from the CvtSk1 field. This does not actually change the monster's mode; only what mode it sounds like. (MonMode.txt for the list of possible inputs).";

            npc["NPC"] = "Open - Points to the matching \"ID\" value in the monstats.txt file. This should not be changed.";
            npc["buy mult"] = "Number - Used to calculate the item's price when it is bought by the NPC from the player. This number is a fraction of 1024 in the following formula: [cost] * [buy mult] / 1024.";
            npc["sell mult"] = "Number - Used to calculate the item's price when it is sold by the NPC to the player. This number is a fraction of 1024 in the following formula: [cost] * [sell mult] / 1024.";
            npc["rep mult"] = "Number - Used to calculate the cost to repair an item. This number is a fraction of 1024 in the following formula: [cost] * [rep mult] / 1024. This is then used to influence the repair cost based on the item durability and charges.";
            npc["questflag A"] = "Number - If the player has this quest flag progress, then apply the relative additional price calculations. Referenced by the Code value of the Quest Flags Table.";
            npc["questbuymult A"] = "Number - Same functionality as the buy mult field, except it applies after it and relies on the questflag A field.";
            npc["questsellmult A"] = "Number - Same functionality as the sell mult field, except it applies after it and relies on the questflag A field.";
            npc["questrepmult A"] = "Number - Same functionality as the rep mult field, except it applies after it and relies on the questflag A field.";
            npc["max buy & max buy (N) & max buy (H)"] = "Number - Sets the maximum price that the NPC will pay, when the player sells an item in Normal Difficulty, Nightmare Difficulty, and Hell Difficulty, respectively";

            objects["Class"] = "Open - Defines the unique type class of the object which is used to reference this object. These are also defined in the objpreset.txt file.";
            objects["Token"] = "Open - Determines what files to use to display the graphics of the object. These are defined by the ObjType.txt file.";
            objects["Selectable0"] = "Boolean - If equals 1, then the object can be selected by the player and highlighted when hovered on by the mouse cursor. If equals 0, then the object cannot be selected and will not highlight when the player hovers the mouse over it.";
            objects["SizeX & SizeY"] = "Number - Controls the amount of sub tiles that the object occupies using X and Y coordinates. This is generally used for measuring the object's size when trying to spawn objects in rooms and controlling their distances apart";
            objects["FrameCnt0"] = "Number - Controls the frame length of the object's mode. If this equals 0, then that mode will be skipped.";
            objects["FrameDelta0"] = "Number - Controls the animation frame rate of how many frames to update per delta (Measured in 256ths).";
            objects["CycleAnim0"] = "Boolean - If equals 1, then the object's current animation will loop back to play again when it finishes. If equals 0, then the object will generally play the Opened mode after playing the Operating mode.";
            objects["Lit0"] = "Number - Controls the Light Radius distance value for the object. If this value equals 0, then the object will not emit a Light Radius.";
            objects["BlocksLight0"] = "Boolean - If equals 1, then the object will draw a shadow. If equals 0, then the object will not draw a shadow.";
            objects["HasCollision0"] = "Boolean - If equals 1, then the object will have collision. If equals 0, then the object will not have collision, and units can walk through it.";
            objects["IsAttackable0"] = "Boolean - If equals 1, then the player can target this object to be attacked, and the player will use the Kick skill when operating the object. If the object has the Class equal to \"CompellingOrb\" or \"SoulStoneForge\", then instead of using the Kick skill, players will use the Attack skill when operating the object. If equals 0, then ignore this, and the player will not use a skill or animation when operating the object.";
            objects["Start0"] = "Number - Controls the frame for where the object will start playing the next animation.";
            objects["EnvEffect"] = "Boolean - If equals 1, then enable the object to update its mode based on the game's time of day. This can mean that when the object is spawned, and it is current day time and the object is in Opened or Operating mode, then it will reset back to Neutral mode. Also, if the current time is dusk, night, or dawn and the object is in Neutral mode, then it will change to Operating mode. If equals 0, then the object will not update its mode based on the time of day.";
            objects["IsDoor"] = "Boolean - If equals 1, then the object will be treated as a door when the game handles its collision, animation properties, tooltips, and commands. If equals 0, then ignore this.";
            objects["BlocksVis"] = "Boolean - If equals 1, then the object will block the player's line of sight to see anything beyond the object. If equals 0, then ignore this. This field relies on the IsDoor field being enabled.";
            objects["Orientation"] = "Number - Determines the object's orientation type, which can affect mouse selection priority of the object when a unit is being rendered in front of or behind the object (such as a door object covering a unit and how the mouse selection should handle that). This also affects the randomization of the coordinates when spawning the object near the edge of a room.";
            objects["OrderFlag0"] = "Number - Controls how the object's Sprite is drawn, which can affect how it is displayed in Perspective game camera mode.";
            objects["PreOperate"] = "Boolean - If equals 1, then enable a random chance that the object will spawn in already in Opened mode. The game will choose a 1/14 chance that this can happen when the object is spawned. If equals 0, then ignore this.";
            objects["Mode0"] = "Boolean - If equals 1, then confirm that this object has the correlating mode. If equals 0, then this object will not have the correlating mode. This flag can affect how the object functions work.";
            objects["Xoffset & Yoffset"] = "Number - Controls the offset values in the X and Y directions for the object's visual graphics. This is measured in game pixels";
            objects["Draw"] = "Boolean - If equals 1, then draw the object's shadows. If equal's 0, then do not draw the object's shadows.";
            objects["Red"] = "Number - Controls the Red color gradient of the object's Light Radius. This field depends on the Lit0 field having a value greater than 0.";
            objects["Green"] = "Number - Controls the Green color gradient of the object's Light Radius. This field depends on the Lit0 field having a value greater than 0.";
            objects["Blue"] = "Number - Controls the Blue color gradient of the object's Light Radius. This field depends on the Lit0 field having a value greater than 0.";
            objects["HD"] = "Boolean - If equals 1, then the object will be flagged to have a Head composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["TR"] = "Boolean - If equals 1, then the object will be flagged to have a Torso composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["LG"] = "Boolean - If equals 1, then the object will be flagged to have a Legs composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["RA"] = "Boolean - If equals 1, then the object will be flagged to have a Right Arm composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["LA"] = "Boolean - If equals 1, then the object will be flagged to have a Left Arm composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["RH"] = "Boolean - If equals 1, then the object will be flagged to have a Right Hand composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["LH"] = "Boolean - If equals 1, then the object will be flagged to have a Left Hand composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["SH"] = "Boolean - If equals 1, then the object will be flagged to have a Shield composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["S1"] = "Boolean - If equals 1, then the object will be flagged to have a Special # composite piece, and the game will use the component system to handle the object's mouse selection collision box. If equals 0, then ignore this.";
            objects["TotalPieces"] = "Number - Defines the total amount of composite pieces. If this value is greater than 1, then the game will treat the object with the multiple composite piece system, and the player can hover the mouse over and select the object's different components.";
            objects["SubClass"] = "Number - Determines the object's class type by declaring a specific value. This is used by the various functions (OperateFn, PopulateFn and InitFn) for knowing how to handle specific types of objects.";
            objects["Xspace & Yspace"] = "Number - Controls the X and Y distance delta values between adjacent objects whe they are being populated together. This field is only used by the PopulateFn field. (Values 3 and 4 for the Add Barrels and Add Crates functions respectively)";
            objects["NameOffset"] = "Number - Controls the vertical offset of the name tooltip's position above the object when the object is being selected. This is measured in pixels.";
            objects["MonsterOK"] = "Boolean - If equals 1, then if a monster operates the object, then the object will run its operate function. If equals 0, then then if a monster operates the object, then the object will not run its operate function.";
            objects["ShrineFunction"] = "Number - Controls what shrine function to use (Code field in Shrines.txt) when the object is told to do its Skill command.";
            objects["Restore"] = "Boolean - If equals 1, the game will restore the object in an inactive state when the area level repopulates after a player loads back into it. If equals 0, then the game will not restore the object.";
            objects["Parm0"] = "Number - Used as possible parameters for various functions for the object.";
            objects["Lockable"] = "Boolean - If equals 1, then the object will have a random chance to spawn with the locked attribute and have a display tooltip name with the \"lockedchest\" [O] This only works when the object has the InitFn value equal to 3. If equals 0, then ignore this.";
            objects["Gore"] = "Number - Controls if an object should call its PopulateFn when it is chosen as an object that can spawn in a room. Objects with a gore value greater than 2 will not be populated in rooms.";
            objects["Sync"] = "Boolean - If equals 1, then the object's animation rate will always match the FrameDelta0 field (depending on the object's mode) which means the client and server will have synced animations. If equals 0, then the animation rate will have random visual variation.";
            objects["Damage"] = "Number - Controls the amount of damage dealt by the object when it performs an OperateFn that deals damage such as triggering a pulse trap or an explosion.";
            objects["Overlay"] = "Boolean - If equals 1, then add and remove an overlay on the object based on its current mode. If equals 0, then ignore this. This field will only work with specific object Classes and will use specific Overlays for those objects.";
            objects["CollisionSubst"] = "Boolean - If equals 1, then the game will handle the bounding box around the object for mouse selection. The game will use the object's pixel size and Left, Top, Width and Height field values to determine the collision size. If equals 0, then ignore this.";
            objects["Left"] = "Number - Controls the starting X position offset value for drawing the bounding collision box around the object for mouse selection. Depends on the CollisionSubst field being enabled.";
            objects["Top"] = "Number - Controls the starting Y position offset value for drawing the bounding collision box around the object for mouse selection. Depends on the CollisionSubst field being enabled.";
            objects["Width"] = "Number - Controls the ending X position offset value for drawing the bounding collision box around the object for mouse selection. Depends on the CollisionSubst field being enabled.";
            objects["Height"] = "Number - Controls the ending Y position offset value for drawing the bounding collision box around the object for mouse selection. Depends on the CollisionSubst field being enabled.";
            objects["OperateFn"] = "Number - Defines a function that the game will use when the player clicks on the object.";
            objects["PopulateFn"] = "Number - Defines a function that the game will use to spawn this object.";
            objects["InitFn"] = "Number - Defines a function to control how the object works while active and when initially activated by a player.";
            objects["ClientFn"] = "Number - Defines a function that runs on the object from the game's client side.";
            objects["RestoreVirgins"] = "Boolean - If equals 1, then when the object has been used, the game will not restore the object in an inactive state when the area level repopulates after a player loads back into it. If equals 0, then ignore this.";
            objects["BlockMissile"] = "Boolean - If equals 1, then missiles can collide with this object. If equals 0, then missiles will ignore and fly through this object.";
            objects["DrawUnder"] = "Number - Controls the targeting priority of the object.";
            objects["OpenWarp"] = "Boolean - If equals 1, then this object will be classified as an object that can be opened to warp to another area, and the UI will be notified to display a tooltip for opening or entering, based on the object's mode. If equals 0, then ignore this.";
            objects["AutoMap"] = "Number - Used to display a tile in the Automap to represent the object. Defines which cell number to use in the tile list for the Automap. If this value equals 0, then this object will not display on the Automap. (Automap.txt).";

            objgroup["GroupName"] = "Open - This is a reference field to define the Object Group name.";
            objgroup["*ID"] = "Number - This field is not read directly, but can be used as an Index for groups";
            objgroup["ID0"] = "Number - Uses the ID field from Objects.txt as an index to choose which Objects are assigned to this Object Group.";
            objgroup["Density0"] = "Number - Controls the number of Objects to spawn in the area level. This is also affected by the Object's populate function defined by the PopulateFn field from the objects.txt file. The maximum value allowed is 128.";
            objgroup["Prob0"] = "Number - Controls the probability that the Object will spawn in the area level. This is calculated in order so the first probability that is successful will be chosen. This also means that these field values should add up to exactly 100 in total to guarantee that one of the objects spawn.";

            objpreset["Index"] = "Number - Assigns a unique numeric ID to the Object Preset so that it can be properly referenced.";
            objpreset["Act"] = "Number - Defines the Act number used for each Object Preset. Uses values between 1 to 5.";
            objpreset["ObjectClass"] = "Open - Uses the Class field from Objects.txt, which assigns an Object to this Object Preset.";

            overlay["overlay"] = "Open - Defines the name of the overlay, used in other data files.";
            overlay["Filename"] = "Open - Defines which DCC file to use for the Overlay.";
            overlay["version"] = "Number - Defines which game version to use this Overlay (0 = Classic mode | 100 = Expansion mode)\"'>.";
            overlay["Character"] = "Open - Used for name categorizing Overlays for unit translation mapping.";
            overlay["PreDraw"] = "Boolean - If equals 1, then display the Overlay in front of Sprites. If equals 0, then display the Overlay behind Sprites.";
            overlay["1ofN"] = "Number - Controls how to randomly display Overlays. This value will randomly add to the current index of the Overlay to possibly use another Overlay that is indexed after this current Overlay. The formula is as follows: Index = Index + RANDOM(0, [\"1ofN\"]-1).";
            overlay["Xoffset"] = "Number - Sets the horizontal offset of the overlay on the unit. Positive values move it toward the left and negative values move it towards the right.";
            overlay["Yoffset"] = "Number - Sets the vertical offset of the overlay on the unit. Positive values move it down and negative values move it up.";
            overlay["Height1"] = "Number - These are additional values added to Yoffset. Only 1 of these Height1 fields are added, and which field gets added depends on the OverlayHeight value from MonStats2.txt (Example: If the \"OverlayHeight\" value is 4, then use the \"Height4\" field). Furthermore, If the \"OverlayHeight\" value is 0, then ignore these \"Height\" fields and add a default value of 75 to \"Yoffset\". Player unit types will always use \"Height2\".";
            overlay["AnimRate"] = "Number - Controls the animation frame rate of the Overlay. The value is the number of frames that will update per second.";
            overlay["LoopWaitTime"] = "Number - Controls the number of periodic frames to wait until redrawing the Overlay. This only works with Overlays that are a loop type.";
            overlay["Trans"] = "Number - Controls the alpha mode for how the Overlay is displayed, which can affect transparency and blending. Referenced by the Code value of the Transparency Table.";
            overlay["InitRadius"] = "Number - Controls the starting Light Radius value for the Overlay (Max = 18).";
            overlay["Radius"] = "Number - Controls the maximum Light Radius value for the Overlay. This can only be greater than or equal to InitRadius. If greater than, the Light Radius will increase in size per frame, starting from InitRadius until it matches this \"Radius\" value (Max = 18).";
            overlay["Red"] = "Number - Controls the Red color gradient of the Light Radius.";
            overlay["Green"] = "Number - Controls the Green color gradient of the Light Radius.";
            overlay["Blue"] = "Number - Controls the Blue color gradient of the Light Radius.";
            overlay["NumDirections"] = "Number - The number of directions in the cell file.";
            overlay["LocalBlood"] = "Number - Controls how to display green blood or VFX on a unit.";

            pettype["PetType"] = "Open - Defines the name of the pet type, used in the PetType column in Skills.txt.";
            pettype["group"] = "Number - Used as an ID field, where if pet types share the same group value, then only 1 pet of that group is allowed to be alive at any time. If equals 0 (or null), then ignore this.";
            pettype["basemax"] = "Number - This sets a baseline maximum number of pets allowed to be alive when skill levels are reset or changed.";
            pettype["warp"] = "Boolean - If equals 1, then the Pet will teleport to the player when the player teleports or warps to another area. If equals 0, then the pet will die instead.";
            pettype["range"] = "Boolean - If equals 1, then the Pet will die if the player teleports or warps to another area and is located more than 40 grid tiles in distances from the Pet. If equals 0, then ignore this.";
            pettype["partysend"] = "Boolean - If equals 1, then tell the Pet to do the Party Location Update command (find the location of its Player) when its health changes. If equals 0, then ignore this.";
            pettype["unsummon"] = "Boolean - If equals 1, then the Pet can be unsummoned by the Unsummon skill function. If equals 0, then the Pet cannot be unsummoned.";
            pettype["automap"] = "Boolean - If equals 1, then display the Pet on the Automap. If equals 0, then hide the pet on the Automap.";
            pettype["drawhp"] = "Boolean - If equals 1, then display the Pet's Life bar under the party frame. If equals 0, then hide the Pet's Life bar under the party icon.";
            pettype["icontype"] = "Number - Controls the functionality for how to display the Pet Icon and number of Pets counter.";
            pettype["baseicon"] = "Open - Define which DC6 file to use for the default Pet's icon in its party frame.";
            pettype["mclass1"] = "Open - Defines the alternative pet to use for the PetType by using that specific unit's hcIDx from Monstats.txt.";
            pettype["micon1"] = "Open - Defines which DC6 file to use for the related \"mclass\" Pet's icon in its party frame.";

            properties["code"] = "Open - Defines the property ID. Used as a reference in other data files.";
            properties["func1"] = "Number - Code function used to define the Property. Uses numeric ID values to define what function to use.";
            properties["stat1"] = "Open - Stat applied by the property. Used by the func1 field as a possible parameter using a Stat entry from ItemStatCost.txt. A stat is comprised of a \"min\" and \"max\" value which it uses to calculate the actual numeric value. Stats also can have a \"parameter\" value, depending on its function.";
            properties["set1"] = "Boolean - Used by the func1 field as a possible parameter. If equals 1, then set the stat value regardless of its current value. If equals 0, then add to the stat value.";
            properties["val1"] = "Number - sed by the func1 field as a possible input parameter for additional function calculations.";

            qualityitems["mod1code"] = "Open - Controls the item properties for the item affix (Uses the Code field from Properties.txt).";
            qualityitems["mod1param"] = "Number - The \"parameter\" value associated with the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            qualityitems["mod1min"] = "Number - The \"min\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            qualityitems["mod1max"] = "Number - The \"max\" value to assign to the listed property (mod). Usage depends on the (Function ID field from Properties.txt).";
            qualityitems["armor"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on both torso armor and helmet item types. If equals 0, then ignore this.";
            qualityitems["weapon"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on melee weapon item types (except scepters, wands, and staves). If equals 0, then ignore this.";
            qualityitems["shield"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on shield item types. If equals 0, then ignore this.";
            qualityitems["scepter"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on scepter item types. If equals 0, then ignore this.";
            qualityitems["wand"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on wand item types. If equals 0, then ignore this.";
            qualityitems["staff"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on staff item types. If equals 0, then ignore this.";
            qualityitems["bow"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on bow or crossbow item types. If equals 0, then ignore this.";
            qualityitems["boots"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on boots item types. If equals 0, then ignore this.";
            qualityitems["gloves"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on gloves item types. If equals 0, then ignore this.";
            qualityitems["belt"] = "Boolean - If equals 1, then allow this High Quality (Superior) modifier to be applied on belt item types. If equals 0, then ignore this.";

            rareprefix["name"] = "Open - Uses a string key to define the Rare Prefix name.";
            rareprefix["version"] = "Number - Defines which game version to use this Set bonus (0 = Classic mode | 100 = Expansion mode).";
            rareprefix["itype1"] = "Open - Controls what item types are allowed for this Rare Prefix to spawn on (Uses the Code field from ItemTypes.txt).";
            rareprefix["etype1"] = "Open - Controls what item types are excluded for this Rare Prefix to spawn on (Uses the Code field from ItemTypes.txt).";

            raresuffix["name"] = "Open - Uses a string key to define the Rare Suffix name.";
            raresuffix["version"] = "Number - Defines which game version to use this Set bonus (0 = Classic mode | 100 = Expansion mode).";
            raresuffix["itype1"] = "Open - Controls what item types are allowed for this Rare Prefix to spawn on (Uses the Code field from ItemTypes.txt).";
            raresuffix["etype1"] = "Open - Controls what item types are excluded for this Rare Prefix to spawn on (Uses the Code field from ItemTypes.txt).";

            runes["Name"] = "Open - Controls the string key that is used to the display the name of the item when the Rune Word is complete.";
            runes["complete"] = "Boolean - If equals 1, then the Rune Word can be crafted in-game. If equals 0, then the Rune Word cannot be crafted in-game.";
            runes["server"] = "Boolean - If equals 1, then the Rune Word can only be crafted on Ladder realm games. If equals 0, then the Rune Word can be crafted in all game types.";
            runes["itype1"] = "Open - Controls what item types are allowed for this Rune Word (Uses the Code field from ItemTypes.txt).";
            runes["etype1"] = "Open - Controls what item types are excluded for this Rune Word (Uses the Code field from ItemTypes.txt).";
            runes["Rune1"] = "Open - Controls what runes are required to make the Rune Word. The order of each of these fields matters. (Uses the Code field from Misc.txt).";
            runes["T1Code1"] = "Open - Controls the item properties that the Rune Word provides (Uses the Code field from Properties.txt).";
            runes["T1Param1"] = "Number - The stat's \"parameter\" value associated with the related property (T1Code1). Usage depends on the (Function ID field from Properties.txt).";
            runes["T1Min1"] = "Number - The stat's \"min\" value to assign to the related property (T1Code1). Usage depends on the (Function ID field from Properties.txt).";
            runes["T1Max1"] = "Number - The stat's \"max\" value to assign to the related property (T1Code1). Usage depends on the (Function ID field from Properties.txt).";

            setitems["index"] = "Open - Links to a string key for displaying the Set item name.";
            setitems["set"] = "Open - Defines the Set to link to this Set Item (must match the Index field from Sets.txt).";
            setitems["item"] = "Open - Defines the baseline item code to use for this Set item (must match the code value from AMW.txt).";
            setitems["rarity"] = "Number - Modifies the chances that this Unique item will spawn compared to the other Set items. This value acts as a numerator and a denominator. Each \"rarity\" value gets summed together to give a total denominator, used for the random roll for the item. For example, if there are 3 possible Set items, and their \"rarity\" values are 3, 5, 7, then their chances to be chosen are 3/15, 5/15, and 7/15 respectively. (The minimum \"rarity\" value equals 1).";
            setitems["lvl"] = "Number - The item level for the item, which controls what object or monster needs to be in order to drop this item.";
            setitems["lvl req"] = "Number - The minimum character level required to equip the item.";
            setitems["chrtransform"] = "Open - Controls the color change of the item when equipped on a character or dropped on the ground. If empty, then the item will have the default item color. Referenced from the Code column in Colors.txt.";
            setitems["invtransform"] = "Open - Controls the color change of the item in the inventory UI. If empty, then the item will have the default item color. Referenced from the Code column in Colors.txt.";
            setitems["invfile"] = "Open - An override for the invfile field from AMW.txt. Uses the item field by default.";
            setitems["flippyfile"] = "Open - An override for the flippyfile field from AMW.txt. Uses the item field by default.";
            setitems["dropsound"] = "Open - An override for the dropsound field from AMW.txt. Uses the item field by default.";
            setitems["dropsfxframe"] = "Number - An override for the dropsfxframe field from AMW.txt. Uses the item field by default.";
            setitems["usesound"] = "Open - An override for the usesound field from the AMW.txt. Uses the item field by default.";
            setitems["cost mult"] = "Number - Multiplicative modifier for the Set item's buy, sell, and repair costs.";
            setitems["cost add"] = "Number - Flat integer modification to the Set item's buy, sell, and repair costs. This is added after the \"cost mult\" has modified the costs.";
            setitems["add func"] = "Number - Controls how the additional Set item properties (aprop#a & aprob#b) will function on the Set item based on other related set items are equipped.";
            setitems["prop1"] = "Open - Controls the item properties that are add baseline to the Set Item (Uses the Code field from Properties.txt).";
            setitems["par1"] = "Number - The stat's \"parameter\" value associated with the related property (prop1). Usage depends on the (Function ID field from Properties.txt).";
            setitems["min1"] = "Number - The stat's \"min\" value to assign to the related property (prop1). Usage depends on the (Function ID field from Properties.txt).";
            setitems["max1"] = "Number - The stat's \"max\" value to assign to the related property (prop1). Usage depends on the (Function ID field from Properties.txt).";
            setitems["aprop1a"] = "Open - Controls the item properties that are added to the Set Item when other pieces of the Set are also equipped (Uses the Code field from Properties.txt).";
            setitems["apar1a"] = "Number - The stat's \"parameter\" value associated with the related property (aprop1a). Usage depends on the (Function ID field from Properties.txt).";
            setitems["amin1a"] = "Number - The stat's \"min\" value to assign to the related property (aprop1a). Usage depends on the (Function ID field from Properties.txt).";
            setitems["amax1a"] = "Number - The stat's \"max\" value to assign to the related property (aprop1a). Usage depends on the (Function ID field from Properties.txt).";
            setitems["aprop1b"] = "Open - Controls the item properties that are added to the Set Item when other pieces of the Set are also equipped. Each of these numbered fields are paired with the related \"aprop#a\" field as an additional item property. (Uses the Code field from Properties.txt).";
            setitems["apar1b"] = "Number - The stat's \"parameter\" value associated with the related property (aprop1b). Usage depends on the (Function ID field from Properties.txt).";
            setitems["amin1b"] = "Number - The stat's \"min\" value to assign to the related property (aprop1b). Usage depends on the (Function ID field from Properties.txt).";
            setitems["amax1b"] = "Number - The stat's \"max\" value to assign to the related property (aprop1b). Usage depends on the (Function ID field from Properties.txt).";
            setitems["diablocloneweight"] = "Number - The amount of weight added to the diablo clone progress when this item is sold. When offline, selling this item will instead immediately spawn diablo clone.";

            sets["index"] = "Open - Defines the specific Set ID.";
            sets["name"] = "Open - Uses a string for displaying the Set name in the inventory tooltip.";
            sets["version"] = "Number - Defines which game version to use this Set bonus (0 = Classic mode | 100 = Expansion mode).";
            sets["PCode2a"] = "Open - Controls the each of the different pairs of Partial Set item properties. These are applied when the player has equipped the related # of Set items. This is the first part of the pair for each Partial Set bonus. (Uses the Code field from Properties.txt).";
            sets["PParam2a"] = "Number - The stat's \"parameter\" value associated with the relative property (PCode#a). Usage depends on the (Function ID field from Properties.txt).";
            sets["PMin2a"] = "Number - The stat's \"min\" value associated with the listed relative (PCode#a). Usage depends on the (Function ID field from Properties.txt).";
            sets["PMax2a"] = "Number - The stat's \"max\" value to assign to the listed relative (PCode#a). Usage depends on the (Function ID field from Properties.txt).";
            sets["PCode2b"] = "Open - Controls the each of the different pairs of Partial Set item properties. These are applied when the player has equipped the related # of Set items. This is the second part of the pair for each Partial Set bonus. (Uses the Code field from Properties.txt).";
            sets["PParam2b"] = "Number - The stat's \"parameter\" value associated with the relative property (PCode#b). Usage depends on the (Function ID field from Properties.txt).";
            sets["PMin2b"] = "Number - The stat's \"min\" value associated with the listed relative (PCode#b). Usage depends on the (Function ID field from Properties.txt).";
            sets["PMax2b"] = "Number - The stat's \"max\" value to assign to the listed relative (PCode#b). Usage depends on the (Function ID field from Properties.txt).";
            sets["FCode1"] = "Open - Controls the each of the different Full Set item properties. These are applied when the player has all Set item pieces equipped (Uses the Code field from Properties.txt).";
            sets["FParam1"] = "Number - The stat's \"parameter\" value associated with the relative property (FCode#b). Usage depends on the (Function ID field from Properties.txt).";
            sets["FMin1"] = "Number - The stat's \"min\" value associated with the listed relative (FCode#b). Usage depends on the (Function ID field from Properties.txt).";
            sets["FMax1"] = "Number - The stat's \"max\" value to assign to the listed relative (FCode#b). Usage depends on the (Function ID field from Properties.txt).";

            shrines["Name"] = "Open - This is a reference field to define the Shrine index.";
            shrines["Code"] = "Number - Code function used to define the Shrine's function. Uses ID values listed below to define what function to use.";
            shrines["Arg0"] = "Number - Integer value used as a possible parameter for the Code function.";
            shrines["Duration in frames"] = "Number - Duration of the effects of the Shrine (Calculated in Frames, where 25 Frames = 1 Second).";
            shrines["reset time in minutes"] = "Number - Controls the amount of time before the Shrine is available to use again. Each value of 1 equals 1200 Frames or 48 seconds. A value of 0 means that the Shrine is a one-time use.";
            shrines["StringName"] = "Open - Uses a string to display as the Shrine's name.";
            shrines["StringPhrase"] = "Open - Uses a string to display as the Shrine's activation phrase when it is used.";
            shrines["effectclass"] = "Number - Used to define the Shrine's archetype which is involved with calculating region stats.";
            shrines["LevelMin"] = "Number - Define the earliest area level where the Shrine can spawn. Area levels are determined from Levels.txt.";

            skills["skill"] = "Open - Defines the unique name ID for the skill, which is how other files can reference the skill. The order of the defined skills will determine their ID numbers, so they should not be reordered.";
            skills["ID"] = "Number - This field is not read directly, but can be used as an Index for skills.";
            skills["charclass"] = "Open - Assigns the skill to a specific character class which affects how skill item modifiers work and what skills the class can learn. Referenced from the Code column in PlayerClass.txt.";
            skills["SkillDesc"] = "Open - Controls the skill's tooltip and general UI display. Points to the SkillDesc field from SkillDesc.txt.";
            skills["srvstfunc"] = "Number - Server Start function. This controls how the skill works when it is starting to cast, on the server side. This uses a code value to call a function, affecting how certain fields are used.";
            skills["srvdofunc"] = "Number - Server Do function. This controls how the skill works when it finishes being cast, on the server side. This uses a code value to call a function, affecting how certain fields are used.";
            skills["srvstopfunc"] = "Number - Server Stop function. This controls how the skill cleans up after ending. This uses a code value to call a function, affecting how certain fields are used. Only used by \"Whirlwind\".";
            skills["prgstack"] = "Boolean - If equals 1, then all srvprgfunc1 functions will execute, based on the current number of progressive charges. If equals 0, then only the relative srvprgfunc1 function will execute, based on the current number of progressive charges.";
            skills["srvprgfunc1"] = "Number - Controls what srvdofunc is used when executing the progressive skill with a charge number equal to 1, 2, and 3, respectively. This field uses the same functions as the srvdofunc field.";
            skills["prgcalc1"] = "Calculation - Used as a possible parameter for calculating values when executing the progressive skill with a charge number equal to 1, 2, and 3, respectively.";
            skills["prgdam"] = "Number - Calls a function to modify the progressive damage dealt.";
            skills["srvmissile"] = "Open - Used as a parameter for controlling what main missile is used for the server functions used (Missile field from Missiles.txt).";
            skills["decquant"] = "Boolean - If equals 1, then the unit's ammo quantity will be updated each time the skill's srvdofunc is called. If equals 0, then ignore this.";
            skills["lob"] = "Boolean - If equals 1, then missiles created by the skill's functions will use the missile lobbing function, which will cause the missile fly in an arc pattern. If equals 0, then missiles that are created will use the normal linear function.";
            skills["srvmissilea"] = "Open - Used as a parameter for controlling what secondary missile is used for the server functions used (Missile field from Missiles.txt).";
            skills["useServerMissilesOnRemoteClients"] = "Broken - Number - Control new missile changes per player skill. Values of 1 will force enable it for that skill. Skills that have matching server/client missiles sets for the skill get auto enabled. Setting this to a value greater than 1 will force it to skip this auto enable logic. If equals 0, then ignore this. Note: this feature is disabled.";
            skills["srvoverlay"] = "Open - Creates an overlay on the target unit when the skill is used. This is a possible parameter used by various skill functions (overlay field from Overlay.txt).";
            skills["aurafilter"] = "Number - Controls different flags that affect how the skill's aura will affect the different types of units. Uses an integer value to check against different bit fields. For example, if the value equals 4354 (binary = 1000100000010) then that returns true for the 4096 (binary = 1000000000000), 256 (binary = 0000100000000), and 2 (binary = 0000000000010) bit field values. (See SuperCalc for a helpful calculator for this).";
            skills["aurastate"] = "Open - Links to a state that can be applied to the caster unit when casting the skill, depending on the skill function used (state field from States.txt).";
            skills["auratargetstate"] = "Open - Links to a state that can be applied to the target unit when using the skill, depending on the skill function used (state field from States.txt).";
            skills["auralencalc"] = "Calculation - Controls the aura state duration on the unit (where 25 Frames = 1 second). If this value is empty, then the state duration will be controlled by other functions, or it will last forever. This can also be used as a parameter for certain skill functions.";
            skills["aurarangecalc"] = "Calculation - Controls the aura state's area radius size, measured in grid sub-tiles. This can also be used as a parameter for certain skill functions.";
            skills["aurastat1"] = "Open - Controls which stat modifiers will be altered or added by the aura state (Stat field from ItemStatCost.txt).";
            skills["aurastatcalc1"] = "Calculation - Controls the value for the relative aurastat1 field.";
            skills["auraevent1"] = "Open - Controls what event will trigger the relative auraeventfunc1 field function. (event field from Events.txt).";
            skills["auraeventfunc1"] = "Number - Controls the function used when the relative auraevent1 event is triggered. Referenced by the Code value of the Event Functions Table.";
            skills["passivestate"] = "Open - Links to a state that can be applied by the passive skill, depending on the skill function used (state field from States.txt).";
            skills["passiveitype"] = "Open - Links to an Item Type to define what type of item needs to be equipped in order to enable the passive state (Code field from Itemtypes.txt).";
            skills["passivereqweaponcount"] = "Number - Controls how many equipped weapons are needed for this passive state to be enabled. If the value equals 1, then the player must have 1 weapon equipped for this passive state to be enabled. If the value equals 2, then the player must be dual wielding weapons for this passive state to be enabled. If the value equals 0, then ignore this field.";
            skills["passivestat1"] = "Open - Assigns stat modifiers to the passive skill (Stat field from ItemStatCost.txt).";
            skills["passivecalc1"] = "Calculation - Controls the value for the relative passivestat1 field.";
            skills["summon"] = "Open - Controls what monster is summoned by the skill (ID field from MonStats.txt). This field's usage will depend on the skill function used. This field can also be used as reference for AI behaviors and for SkillDesc.txt.";
            skills["PetType"] = "Open - Links to a pet type data to control how the summoned unit is displayed on the UI (PetType column in PetType.txt).";
            skills["petmax"] = "Calculation - Used skill functions that summon pets to control how many summon units are allowed at a time.";
            skills["summode"] = "Open - Defines the animation mode that the summoned monster will be initiated with. Referenced from the Code column in MonMode.txt.";
            skills["sumskill1"] = "Open - Assigns a skill to the summoned monster. Points to another Skill. This can be useful for adding a skill to a monster to transition its synergy bonuses.";
            skills["sumsk1calc"] = "Calculation - Controls the skill level for the designated sumskill1 field when the skill is assigned to the monster.";
            skills["sumumod"] = "Number - Assigns a monster modifier to the summoned monster (ID field from MonUMod.txt).";
            skills["sumoverlay"] = "Open - Creates an overlay on the summoned monster when it is first created (overlay field from Overlay.txt).";
            skills["stsuccessonly"] = "Boolean - If equals 1, then the following sound and overlay fields will only play when the skill is successfully cast, instead of always being used even when the skill cast is interrupted. If equals 0, then the following sound and overlay fields will always be used when the skill is cast, regardless if the skill was interrupted or not.";
            skills["stsound"] = "Open - Controls what client sound is played when the skill is used, based on the client starting function (Sound field from Sounds.txt).";
            skills["stsoundclass"] = "Open - Controls what client sound is played when the skill is used by the skill's assigned class (charclass), based on the client starting function (Sound field from Sounds.txt). If the unit using the skill is not the same class as the \"charclass\" value for the skill, then this sound will not play.";
            skills["stsounddelay"] = "Boolean - If equals 1, then use the weapon's hit class to determine the delay in frames (where 25 frames = 1 second) before playing the skill's start sound. If equals 0, then the skill's start sound will play immediately.";
            skills["weaponsnd"] = "Boolean - If equals 1, then play the weapon's hit sound when hitting an enemy with this skill. The sound chosen is based on the weapon's hit class. Also use the sound delay based on the weapon's hit class to determine the delay in frames (where 25 frames = 1 second) before playing the weapon hit sound (stsounddelay for the types of hit class sounds and delays used). If equals 0, then do not play the weapon hit sound when hitting an enemy with the skill attack.";
            skills["dosound"] = "Open - Controls the sound for the skill each time the cltdofunc is used (Sound field from Sounds.txt).";
            skills["dosound a"] = "Open - Used as a possible parameter for playing additional sounds based on the cltdofunc used(Sound field from Sounds.txt).";
            skills["tgtoverlay"] = "Open - Used as a possible parameter for adding an Overlay on the target unit, based on the cltdofunc used (overlay field from Overlay.txt).";
            skills["tgtsound"] = "Open - Used as a possible parameter for playing a sound located on the target unit, based on the cltdofunc used (Sound field from Sounds.txt).";
            skills["prgoverlay"] = "Open - Used as a possible parameter for adding an Overlay on the caster unit for progressive charge-up skill functions, based on the cltdofunc used and how many progressive charges the caster unit has (overlay field from Overlay.txt).";
            skills["prgsound"] = "Open - Used as a possible parameter for playing a sound when using the skill for progressive charge-up skill functions, based on the cltdofunc used and how many progressive charges the caster unit has (Sound field from Sounds.txt).";
            skills["castoverlay"] = "Open - Used as a possible parameter for adding an Overlay on the caster unit when using the skill, based on the Client Start/Do function used (overlay field from Overlay.txt).";
            skills["cltoverlaya"] = "Open - Used as a possible parameter for adding additional Overlays on the caster unit, based on the Client Start/Do function used (overlay field from Overlay.txt).";
            skills["cltstfunc"] = "Number - Client Start function. This controls how the skill works when it is starting to cast, on the client side. This uses a code value to call a function, affecting how certain fields are used.";
            skills["cltdofunc"] = "Number - Client Do function. This controls how the skill works when it finishes being cast, on the client side. This uses a code value to call a function, affecting how certain fields are used.";
            skills["cltstopfunc"] = "Number - Client Stop function. This controls how the skill cleans up after ending. This uses a code value to call a function, affecting how certain fields are used. Only used for \"Whirlwind\".";
            skills["cltprgfunc1"] = "Number - Controls which Client Do function is used when the skill is executed while having a progressive charge number equal to 1, 2, and 3, respectively. (uses cltdofunc values).";
            skills["cltmissile"] = "Open - Used as a parameter for controlling what main missile is used for the client functions used (Missile field from Missiles.txt).";
            skills["cltmissilea"] = "Open - Used as a parameter for controlling what secondary missile is used for the client functions used (Missile field from Missiles.txt).";
            skills["cltcalc1"] = "Calculation - Use as a possible parameter for controlling values for the client functions.";
            skills["warp"] = "Boolean - If equals 1 and the skill uses a function that involves warping/teleporting, then check for a scene transition loading screen, if necessary. If equals 0, then ignore this.";
            skills["immediate"] = "Boolean - If equals 1 and the skill has a periodic function, then immediately perform the skill's function when the periodic skill is activated. If equals 0, then wait until the skill's first periodic delay before performing the skill's function.";
            skills["enhanceable"] = "Boolean - If equals 1, then the skill will be included in the plus to all skills item modifier. If equals 0, then the skill will not be included in the plus to all skills item modifier.";
            skills["noammo"] = "Boolean - If equals 1, then the skill will not check that weapon's ammo when performing an attack. This relies on the Shoots field from ItemTypes.txt. If equals 0, then the weapon will consume its ammo when being used by the skill.";
            skills["weapsel"] = "Number - If the unit can dual wield weapons, then this field will control how the weapons are used for the skill.";
            skills["itypea1"] = "Open - Controls what Item Types are included, or allowed, when determining if this skill can be used (Code field from ItemTypes.txt).";
            skills["etypea1"] = "Open - Controls what Item Types are excluded, or not allowed, when determining if this skill can be used (Code field from ItemTypes.txt).";
            skills["itypeb1"] = "Open - Controls what alternate Item Types are included, or allowed, when determining if this skill can be used (Code field from ItemTypes.txt). This acts as a second set of Item Types to check.";
            skills["etypeb1"] = "Open - Controls what alternate Item Types are excluded, or not allowed, when determining if this skill can be used (Code field from ItemTypes.txt). This acts as a second set of Item Types to check.";
            skills["anim"] = "Open - Controls the animation mode that the player character will use when using this skill. Setting the mode to Sequence (SQ) will cause the player character to play a time controlled animation sequence, utilizing certain sequence fields. Referenced from the Token column in PlrMode.txt.";
            skills["seqtrans"] = "Open - Uses the same inputs as the anim field. If the \"anim\" field equals SQ (Sequence) and this field equals SC (Cast), then the sequence animation speed can be modified by the faster cast rate stat modifier.";
            skills["monanim"] = "Open - Controls the animat on mode that the monster will use when using this skill. This is similar to the anim field except with monster units using the skill instead of player units. Referenced from the Token column in MonMode.txt.";
            skills["seqnum"] = "Number - If the unit is a player and the anim used for the skill is a Sequence, then this field will control the index of which sequence animation to use. These sequences are specifically designed for certain skills, and each sequence has a set number of frames using certain animations.";
            skills["seqinput"] = "Number - For skills that can repeat, this controls the number of frames to wait before the \"Do\" frame in the sequence. This acts as a delay in frames (25 Frames = 1 second) to wait within the sequence animation before it is allowed to be cast again or for looping back to the start of the sequence, such as for the Sorceress Inferno skill.";
            skills["durability"] = "Boolean - If equals 1 and when the player character ends an animation mode that was not Attack 1, Attack 2, or Throw, then check the quantity and durability of the player's items to see if the valid weapons are out of ammo or are broken. If equals 0, then ignore this.";
            skills["UseAttackRate"] = "Boolean - If equals 1, then the current attack speed should be updated after using the skill, relative to the \"attackrate\" stat (ItemStatCost.txt), and if the skill was an attacking skill. If equals 0, then the attack speed will not be updated after using the skill.";
            skills["LineOfSight"] = "Number - Controls the skill's collision filters for valid target locations when it is being cast.";
            skills["TargetableOnly"] = "Boolean - If equals 1, then the skill will require a target unit in order to be used. If equals 0, then ignore this.";
            skills["SearchEnemyXY"] = "Boolean - If equals 1, then when the skill is cast on a target location, it will automatically search in different directions in the target area to find the first available enemy target. If equals 0, then ignore this. This field can be overridden if the SearchEnemyNear field is enabled.";
            skills["SearchEnemyNear"] = "Boolean - If equals 1, then when the skill is cast on a target location, it will automatically find the nearest enemy target. If equals 0, then ignore this.";
            skills["SearchOpenXY"] = "Boolean - If equals 1, then automatically search in a radius of size 7 in around the clicked target location to find an available unoccupied location to use the skill, testing for collision. If equals 0, then ignore this. This field can be overridden if the SearchEnemyNear or SearchEnemyXY field is enabled.";
            skills["SelectProc"] = "Number - Uses a function to check that the target is a corpse with various parameters before validating the usage of the skill.";
            skills["TargetCorpse"] = "Boolean - If equals 1, then the skill is allowed to target corpses. If equals 0, then skill cannot target corpses.";
            skills["TargetPet"] = "Boolean - If equals 1, then the skill is allowed to target pets (summons and mercenaries). If equals 0, then the skill cannot target pets.";
            skills["TargetAlly"] = "Boolean - If equals 1, then the skill is allowed to target ally units. If equals 0, then the skill cannot target ally units.";
            skills["TargetItem"] = "Boolean - If equals 1, then the skill is allowed to target items on the ground. If equals 0, then the skill cannot target items.";
            skills["AttackNoMana"] = "Boolean - If equals 1, then then when the caster does not have enough mana to cast the skill and attempts to use the skill, the caster will default to using the Attack skill. If equals 0, then attempting to use the skill without enough mana will do nothing.";
            skills["TgtPlaceCheck"] = "Boolean - If equals 1, then check that the target location is available for spawning a unit, testing collision. If equals 0, then ignore this. This is only used for skills that monsters use to spawn new units.";
            skills["KeepCursorStateOnKill"] = "Boolean - If equals 1, then the mouse click hold state will continue to stay active after killing a selected target. If equals 0, then after killing the target, the mouse click hold state will reset.";
            skills["ContinueCastUnselected"] = "Boolean - If equals 1, then while the mouse is held down and there is no more target selected, then the skill will continue being used at the mouse's location. If equals 0, then while the mouse is held down and there is no more target selected, then the player character will cancel the skill and move to the mouse location.";
            skills["ClearSelectedOnHold"] = "Boolean - If equals 1, then when the mouse is held down, the target selection will be cleared. If equals 0, then ignore this.";
            skills["ItemEffect"] = "Number - Uses a Server Do function (\"srvdofunc\") for handling how the skill is used when it is triggered by an item, on the server side.";
            skills["ItemCltEffect"] = "Number - Uses a Client Do function (\"cltdofunc\") for handling how the skill is used when it is triggered by an item, on the client side.";
            skills["ItemTgtDo"] = "Boolean - If equals 1, then use the skill from the enemy target instead of the caster. If equals 0, then ignore this.";
            skills["ItemTarget"] = "Number - Controls the targeting behavior of the skill when it is triggered by an item.";
            skills["ItemCheckStart"] = "Boolean - If equals 1, then use the skill's srvstfunc when the skill is trigged by an item. If equals 0, then the skill's Server Start function is ignored.";
            skills["ItemCltCheckStart"] = "Boolean - If equals 1, then when the skill is triggered by an item, and if the target is dead and the skill has a cltstfunc, then add the \"corpse_noselect\" to the target. If equals 0, then ignore this.";
            skills["ItemCastSound"] = "Open - Play a sound when the skill is used by an item event. (Sound field from Sounds.txt).";
            skills["ItemCastOverlay"] = "Open - Add a cast overlay when the skill is used by an item event. (overlay field from Overlay.txt).";
            skills["skpoints"] = "Calculation - Controls the number of Skill Points needed to level up the skill.";
            skills["reqlevel"] = "Number - Minimum character level required to be allowed to spend Skill Points on this skill.";
            skills["maxlvl"] = "Number - Maximum base level for the skill from spending Skill Points. Skill levels can be increased beyond this through item modifiers.";
            skills["reqstr"] = "Number - Minimum Strength attribute required to be allowed to spend Skill Points on this skill.";
            skills["reqdex"] = "Number - Minimum Dexterity attribute required to be allowed to spend Skill Points on this skill.";
            skills["reqint"] = "Number - Minimum Intelligence attribute required to be allowed to spend Skill Points on this skill.";
            skills["reqvit"] = "Number - Minimum Vitality attribute required to be allowed to spend Skill Points on this skill.";
            skills["reqskill1"] = "Open - Points to a Skill field to act as a prerequisite skill. The prerequisite skill must be least base level 1 before the player is allowed to spend Skill Points on this skill.";
            skills["restrict"] = "Number - Controls how the skill is used when the unit has a restricted state (restrict field from States.txt).";
            skills["State1"] = "Open - Points to a state field from States.txt. Used as parameters for the restrict field to control what specific States will restrict the usage of the skill.";
            skills["localdelay"] = "Calculation - Controls the Casting Delay duration for this skill after it is used (25 frames = 1 second).";
            skills["globaldelay"] = "Calculation - Controls the Casting Delay duration for all other skills with Casting Delays after this skill is used (25 frames = 1 second).";
            skills["leftskill"] = "Boolean - If equals 1, then the skill will appear on the list of skills to assign for the Left Mouse Button. If equals 0, then the skill will not appear on the Left Mouse Button skill list.";
            skills["rightskill"] = "Boolean - If equals 1, then the skill will appear on the list of skills to assign for the Right Mouse Button. If equals 0, then the skill will not appear on the Right Mouse Button skill list.";
            skills["repeat"] = "Boolean - If equals 1 and the skill has no \"anim\" mode declared, then on the client side, the unit's mode will repeat its current mode (this can also happen if the unit has the \"inferno\" state). If equals 0, then the unit will have its mode set to Neutral when starting to use the skill.";
            skills["alwayshit"] = "Boolean - If equals 1, then skills that rely on attack rating and defense will ignore those stats and will always hit enemies. If equals 0, then ignore this.";
            skills["usemanaondo"] = "Boolean - If equals 1, then the skill will consume mana on its do function instead of its start function. If equals 0, then the skill will consume mana on its start function, which is the general case for skills.";
            skills["startmana"] = "Number - Controls the required amount of mana to start using the skill. This can control skills that reduce in mana cost per level to not fall below this amount. (srvst routine).";
            skills["minmana"] = "Number - Controls the minimum amount of mana to use the skill. This can control skills that reduce in mana cost per level to not fall below this amount. (srvdo routine).";
            skills["manashift"] = "Number - This acts as a multiplicative modifier to control the precision of the mana cost after calculating the total mana cost with the mana and lvlmana fields. Mana is calculated in 256ths in code which equals 8 bits. This field acts as a bitwise shift value, which means it will modify the mana value by the power of 2. For example, if this value equals 5 then that means divide the mana value of 256ths by 2^5 = 32 (or multiply the mana by 32/256). A value equal to 8 means 256/256 which means that the mana of 256ths value is not shifted.";
            skills["mana"] = "Number - Defines the base mana cost to use the skill at level 1.";
            skills["lvlmana"] = "Number - Defines the change in mana cost per skill level gained.";
            skills["interrupt"] = "Boolean - If equals 1, then the casting the skill will be interruptible such as when the player is getting hit while casting a skill. If equals 0, then the skill should be uninterruptible.";
            skills["InTown"] = "Boolean - If equals 1, then the skill can be used while the unit is in town. If equals 0, then the skill cannot be used in town.";
            skills["aura"] = "Boolean - If equals 1, then the skill will be classified as an aura, which will make the skill execute its function periodically (perdelay), based on the if the aurastate state is active. Aura skills will also override a preexisting state if that state matches the skill's \"aurastate\" state. If equals 0, then ignore this.";
            skills["periodic"] = "Boolean - If equals 1, then the skill will execute its function periodically (perdelay), based on the if the aurastate state is active. If equals 0, then ignore this.";
            skills["perdelay"] = "Calculation - Controls the periodic rate that the skill continuously executes its function. Minimum value equals 5. This field requires periodic or aura field to be enabled.";
            skills["finishing"] = "Boolean - If equals 1, then the skill will be classified as a finishing move, which can affect how progressive charges are consumed when using the skill and how the skill's description tooltip is displayed. If equals 0, then ignore this.";
            skills["prgchargestocast"] = "Number - Controls how many progressive charges are required to cast the skill.";
            skills["prgchargesconsumed"] = "Number - Controls how many progressive charges are consumed when the skill attack hits an enemy.";
            skills["passive"] = "Boolean - If equals 1, then the skill will be treated as a passive, which can mean that the skill will not show up in the skill selection menu and will not run a server do function. If equals 0, then the skill is an active skill that can be used.";
            skills["progressive"] = "Boolean - If equals 1, then the skill will use the progressive calculations to act as a charge-up skill that will add charges. This is only used for certain skill functions and will generally require the usage of the prgcalc1 and aurastat fields. If equals 0, then ignore this.";
            skills["scroll"] = "Boolean - If equals 1, then the skill can appear as a scroll version in the skill selection UI, if the skill allows for the scroll mechanics and if the player has the skill's scroll item in the inventory. If equals 0, then ignore this.";
            skills["calc1"] = "Calculation - It is used as a possible parameter for skill functions or as a numeric input for other calculation fields.";
            skills["Param1"] = "Number - It is used as a possible parameter for skill functions or as a numeric input for other calculation fields.";
            skills["InGame"] = "Boolean - If equals 1, then the skill is enabled to be used in-game. If equals 0, then the skill will be disabled in-game.";
            skills["ToHit"] = "Number - Baseline bonus Attack Rating that is added to the attack when using this skill at level 1.";
            skills["LevToHit"] = "Number - Additional bonus Attack Rating when using this skill, calculated per skill level.";
            skills["ToHitCalc"] = "Calculation - Calculates the bonus Attack Rating when using the skill. This will override the ToHit and LevToHit fields if it is not blank.";
            skills["ResultFlags"] = "Number - Controls different flags that can affect how the target reacts after being hit by the missile. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 5 (binary = 101) then that returns true for both the 4 (binary = 100) and 1 (binary = 1) bit field values. Referenced by the Bit Field Value of the Result Flags Table.";
            skills["HitFlags"] = "Number - Controls different flags that can affect the damage dealt when the target is hit by the missile. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 6 (binary = 110) then that returns true for both the 4 (binary = 100) and 2 (binary = 10) bit field values. Referenced by the Bit Field Value of the Hit Flags Table.";
            skills["HitClass"] = "Number - Defines the skill's damage routines when hitting, mainly used for determining hit sound effects and overlays. Uses an integer value to check against different bit fields by using the \"&\" operator. For example, if the value equals 6 (binary = 110) then that returns true for both the 4 (binary = 100) and 2 (binary = 10) bit field values. There are binary masks to check between Base Hit Classes and Hit Class Layers, which can distinguish bewteen overlays or sounds are displayed.";
            skills["Kick"] = "Boolean - If equals 1, then a separate function is used to calculate the physical damage dealt by the skill, ignoring the following damage fields. This function will factor in the player character's Strength and Dexterity attributes (or Monster's level) to determine the baseline damage dealt. If equals 0, then ignore this.";
            skills["HitShift"] = "Number - Controls the percentage modifier for the skill's damage. This value cannot be less than 0 or greater than 8. This is calculated in 256ths.";
            skills["SrcDam"] = "Number - Controls the percentage modifier for how much weapon damage is transferred to the skill's damage (Out of 128). For example, if the value equals 64, then 50% (64/128) of the weapon's damage will be added to the skill's damage.";
            skills["MinDam"] = "Number - Minimum baseline physical damage dealt by the skill.";
            skills["MinLevDam1"] = "Number - Controls the additional minimum physical damage dealt by the skill, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field number.";
            skills["MaxDam"] = "Number - Maximum baseline physical damage dealt by the skill.";
            skills["MaxLevDam1"] = "Number - Controls the additional maximum physical damage dealt by the skill, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field number.";
            skills["DmgSymPerCalc"] = "Calculation - Determines the percentage increase to the physical damage dealt by the skill.";
            skills["EType"] = "Open - Defines the type of elemental damage dealt by the skill. If this field is empty, then the related elemental fields below will not be used. Referenced by the Code value of the Elemental Types Table.";
            skills["EMin"] = "Number - Minimum baseline elemental damage dealt by the skill.";
            skills["EMinLev1"] = "Number - Controls the additional minimum elemental damage dealt by the skill, calculated using the leveling formula between 5 level thresholds of the skill's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field number.";
            skills["EMax"] = "Number - Maximum baseline elemental damage dealt by the skill.";
            skills["EMaxLev1"] = "Number - Controls the additional maximum elemental damage dealt by the skill, calculated using the leveling formula between 5 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17-22, 23-28, 29 and beyond. These 5 level thresholds correlate to each field.";
            skills["EDmgSymPerCalc"] = "Calculation - Determines the percentage increase to the total elemental damage dealt by the skill.";
            skills["ELen"] = "Number - The baseline elemental duration dealt by the skill. This is calculated in frame lengths where 25 Frames = 1 second. These fields only apply to appropriate elemental types with a duration.";
            skills["ELevLen1"] = "Number - Controls the additional elemental duration added by the skill, calculated using the leveling formula between 3 level thresholds of the missile's current level. The level thresholds are levels 2-8, 9-16, 17 and beyond. These 3 level thresholds correlate to each field number. These fields only apply to appropriate elemental types with a duration.";
            skills["ELenSymPerCalc"] = "Calculation - Determines the percentage increase to the total elemental duration dealt by the skill.";
            skills["aitype"] = "Number - Controls what the skill's archetype for how the AI will handle using this skill. This mostly affects the mercenary AI and Shadow Warrior AI, but some types are used for general AI.";
            skills["aibonus"] = "Number - This is only used with the Shadow Master AI. This value is a flat integer rating that gets added to the AI's parameters when deciding which skill is most likely to be used next. The higher this value, then the more likely this skill will be used by the AI.";
            skills["cost mult"] = "Number - Multiplicative modifier of an item's gold cost, only when the item has a stat modifier with this skill. This will affect the item's buy, sell, and repair costs (Calculated in 1024ths).";
            skills["cost add"] = "Number - Flat integer modification of an item's gold cost, only when the item has a stat modifier with this skill. This will affect the item's buy, sell, and repair costs. This is added after the cost mult has modified the costs.";

            skilldesc["SkillDesc"] = "Open - The name of the skill description, as a reference for associated Data files.";
            skilldesc["SkillPage"] = "Number - Determines which page on the Skill tree to display the skill.";
            skilldesc["SkillRow"] = "Number - Determines which row on the Skill tree page to display the skill.";
            skilldesc["SkillColumn"] = "Number - Determines which column on the Skill tree page to display the skill.";
            skilldesc["ListRow"] = "Number - Determines which row the skill will be listed in, for the skill select UI.";
            skilldesc["IconCel"] = "Number - Determines the icon asset for displaying the skill, using the Frame# of the Sprite (See Common Folders). It will use Frame+1 for the \"when pressed\" icon visual. The sprite path can be changed in _ProfileHD.json or _ProfileLV.json.";
            skilldesc["HireableIconCel"] = "Number - Determines the icon asset for displaying the skill, using the Frame# of the Sprite (See Common Folders). There is no \"when pressed\" icon visual. The sprite path can be changed in _ProfileHD.json or _ProfileLV.json.";
            skilldesc["str name"] = "Open - Uses a string to display as the skill name.";
            skilldesc["str short"] = "Open - Uses a string to display as the skill description in shortcuts or when selecting a skill.";
            skilldesc["str long"] = "Open - Uses a string to display as the skill description on the Skill Tree.";
            skilldesc["str alt"] = "Open - Uses a string to display the skill name on the Character Screen when the skill is selected.";
            skilldesc["descdam"] = "Number - Use a function to calculate a skill's damage and determine how to display it. These functions sometimes require certain skill fields, especially the damage related fields.";
            skilldesc["ddam calc1"] = "Calculation - Integer calc value used as a possible parameter for the descdam function.";
            skilldesc["p1dmelem"] = "Open - Used for skills that have charge-ups to display the damage on the Character Screen, controls the elemental type for that charge.";
            skilldesc["p1dmmin"] = "Calculation - Used for skills that have charge-ups to display the damage on the Character Screen, controls the minimum damage for that charge.";
            skilldesc["p1dmmax"] = "Calculation - Used for skills that have charge-ups to display the damage on the Character Screen, controls the maximum damage for that charge.";
            skilldesc["descatt"] = "Number - Used to display the overall Attack Rating of the skill in the Character Screen.";
            skilldesc["descmissile1"] = "Open - Links a missile from Missiles.txt to be used as a reference value for calculations.";
            skilldesc["descline1"] = "Number - Uses an ID value to select a description function to format the string value. Displays this text as the current level and next level description lines in the skill tooltip.";
            skilldesc["desctexta1"] = "Open - String value used as the first possible string parameter for the descline1 function.";
            skilldesc["desctextb1"] = "Open - String value used as the second possible string parameter for the descline1 function.";
            skilldesc["desccalca1"] = "Calculation - Integer calculation value used as the first possible numeric parameter for the descline1 function.";
            skilldesc["desccalcb1"] = "Calculation - Integer calculation value used as the second possible numeric parameter for the descline1 function.";
            skilldesc["dsc2line1"] = "Number - Uses an ID value to select a description function to format the string value. Displays this text as a pinned line, after the skill description. (Uses the same function codes as descline1).";
            skilldesc["dsc2texta1"] = "Open - String value used as the first possible string parameter for the dsc2line1 function.";
            skilldesc["dsc2textb1"] = "Open - String value used as the second possible string parameter for the dsc2line1 function.";
            skilldesc["dsc2calca1"] = "Calculation - Integer Calc value used as the first possible numeric parameter for the dsc2line1 function.";
            skilldesc["dsc2calcb1"] = "Calculation - Integer Calc value used as the second possible numeric parameter for the dsc2line1 function.";
            skilldesc["dsc3line1"] = "Number - Uses an ID value to select a description function to format the string value. Displays this text as a pinned line at the bottom of the skill tooltip. (Uses the same function codes as descline1).";
            skilldesc["dsc3texta1"] = "Open - String value used as the first possible string parameter for the dsc3line1 function.";
            skilldesc["dsc3textb1"] = "Open - String value used as the second possible string parameter for the dsc3line1 function.";
            skilldesc["dsc3calca1"] = "Calculation - Integer Calc value used as the first possible numeric parameter for the dsc3line1 function.";
            skilldesc["dsc3calcb1"] = "Calculation - Integer Calc value used as the second possible numeric parameter for the dsc3line1 function.";

            sounds["Sound"] = "Open - Defines the unique name ID for the sound, which is how other files can reference the sound.";
            sounds["Redirect"] = "Open - Points the sound so the index of another sound in the data file. If this field is not empty, the game will use the redirected sound instead of this sound. This can be used when playing the game in the new graphics mode.";
            sounds["Channel"] = "Open - Declares which channel the sound is initialized in. This can affect how different volume or sound settings handle this sound.";
            sounds["FileName"] = "Open - Defines the file path and name of the sound file to play.";
            sounds["IsLocal"] = "Boolean - If equals 1, then this sound is considered a localized sound and will change based on the game's localization setting. If equals 0, then ignore this.";
            sounds["IsMusic"] = "Boolean - If equals 1, then the sound is flagged as a music sound, which affects how music related settings handle this sound. If equals 0, then ignore this.";
            sounds["IsAmbientScene"] = "Boolean - If equals 1, then the sound is flagged as an ambient scene sound, which affects how the game handles the sound when the player transitions between areas. If equals 0, then ignore this.";
            sounds["IsAmbientEvent"] = "Boolean - If equals 1, then the sound is flagged as an ambient event sound, which affects how the game treats the sound when the player transitions between areas. If equals 0, then ignore this.";
            sounds["IsUI"] = "Boolean - If equals 1, then the sound is flagged as a UI sound, which affects how UI related settings handle this sound. If equals 0, then ignore this.";
            sounds["Volume Min"] = "Number - Controls the minimum volume of the sound. Uses a range of 0 to 255.";
            sounds["Volume Max"] = "Number - Controls the maximum volume of the sound. If both Volume Min and Volume Max fields differ in value, then the sound will randomly select a volume value in between these values when it is played. Uses a range of 0 to 255.";
            sounds["Pitch Min"] = "Number - Controls the minimum pitch percentage of the sound.";
            sounds["Pitch Max"] = "Number - Controls the maximum pitch percentage of the sound. If both Pitch Min and Pitch Max fields differ in value, then the sound will randomly select a pitch value in between these values when it is played.";
            sounds["Group Size"] = "Number - Defines a sound Group by declaring a size value. When the sound has this value greater than 0, then this sound is declared as the group's base sound. Any link to use a sound should use the base sound, to signify that the game should use this group of sounds. This field's value controls the number of sounds indexed after base sound that should be added to the group. For example, if the sound has a \"Group Size\" value equal to 5, then this sound is declared as the group's base sound, and the next 4 sounds indexed after this base sound will be added to the group.";
            sounds["Group Weight"] = "Number - Controls the chance to pick the sound when it is part of a group with other sounds. If all sounds in the group do not have a \"Group Weight\" value, then the group sounds will play in historical order. This value controls a weighted random chance, meaning that all related sounds have their weights added together for a total chance and each sound's weight value is rolled against that total value to determine if the sound is successfully picked. The higher this value, the more likely the sound will be picked. This is only used when the sound is part of a group (\"Group Size\").";
            sounds["Loop"] = "Boolean - If equals 1, then the sound will replay itself after it finishes playing. If equals 0, then the sound will only play once.";
            sounds["Fade In"] = "Number - Controls how long to gradually increase the sound's volume starting from 0 when the sound starts playing. Measured in audio game ticks, where 1 game frame is 40 audio ticks, and the game runs at 25 frames per second.";
            sounds["Fade Out"] = "Number - Controls how long to gradually decrease the sound's volume to 0 when the sound stops playing. Measured in audio game ticks, where 1 game frame is 40 audio ticks, and the game runs at 25 frames per second.";
            sounds["Defer Inst"] = "Boolean - If equals 1, then when a duplicate instance of this sound plays the game will stop that request. If equals 0, then ignore this.";
            sounds["Stop Inst"] = "Boolean - If equals 1, then when a duplicate instance of this sound plays the previous instance of the sound will stop and the new instance of the sound will play. If equals 0, then ignore this.";
            sounds["Duration"] = "Number - Controls the length of time to play the sound. When the sound has been playing for this length of time, then the sound will stop. If this equals 0, then ignore this functionality.";
            sounds["Compound"] = "Number - Controls the game tick time limit for when a sound can join in playing based on the previous sound played in the Group. If equals 0, then the sound will not be compounded.";
            sounds["Falloff"] = "Number - Defines the range of falloff for hearing the sound, based on distance. Uses a code to determine the range value presets.";
            sounds["LFEMix"] = "Number - Controls the percentage (out of 100) of the sound's Low-Frequency Effects channel.";
            sounds["3dSpread"] = "Number - Controls the 3D spread angle of the sound. This only works if the sound is considered a 3D sound (Is2D).";
            sounds["Priority"] = "Number - Controls which if the sound should play before other sounds when too many sounds are playing at once. This value is compared to the priority value of other sounds, and the sound that has the higher priority will play first. Sounds belonging to the player will get an increased priority value of 80.";
            sounds["Stream"] = "Boolean - If equals 1, then the sound will be file streamed into the game when called to play. If equals 0, then the entire sound will be loaded into the game before playing.";
            sounds["Is2D"] = "Boolean - If equals 1, then the sound is considered a 2D sound and will not have 3D spread settings. If equals 0, then the sound is considered a 3D sound and will use the 3D spread settings.";
            sounds["Tracking"] = "Boolean - If equals 1, then the sound will track a unit and will update its position to follow that unit. If equals 0, then the sound will not move and will be stationary.";
            sounds["Solo"] = "Boolean - If equals 1, then reduce the volume of other sounds while this sound is playing. If equals 0, then ignore this.";
            sounds["Music Vol"] = "Boolean - If equals 1, then the sound's volume will be affected by the music volume in the game options menu. If equals 0, then ignore this.";
            sounds["Block 1"] = "Number - Defines an offset time value in the sound. If this sound is used in a Sound Environment then these fields control when to periodically update the current song sound to an offset. If this sound is not used in a Sound Environment and Block 1 is used and the Loop field is enabled, then use this block value as the time in the sound when to start looping. If this equals -1, then the field is ignored.";
            sounds["HDOptOut"] = "Boolean - If equals 1, then the sound will not play in the new graphics mode. If equals 0, then the sound will play in the new graphics mode.";
            sounds["Delay"] = "Number - Adds a delay to the starting tick of the sound when the sound starts playing. Measured in audio game ticks, where 1 game frame is 40 audio ticks, and the game runs at 25 frames per second.";

            soundenviron["Handle"] = "Open - A reference field to define the name of the Sound Environment.";
            soundenviron["Index"] = "Open - A reference field to define the ID/Index for the Environment.";
            soundenviron["Song"] = "Open - Play this sound as the background music while the player is in an area level.";
            soundenviron["Day Ambience"] = "Open - Play this sound as an ambient sound while it is currently daytime in the game.";
            soundenviron["HD Day Ambience"] = "Open - Play this sound as an ambient sound while it is currently daytime in the game while playing in the new graphics mode.";
            soundenviron["Night Ambience"] = "Open - Play this sound as an ambient sound while it is currently nighttime in the game.";
            soundenviron["HD Night Ambience"] = "Open - Play this sound as an ambient sound while it is currently nighttime in the game while playing in the new graphics mode.";
            soundenviron["Day Event"] = "Open - Play this sound at a random range and variance in the background when it is currently daytime in the game.";
            soundenviron["HD Day Event"] = "Open - Play this sound at a random range and variance in the background when it is currently daytime in the game while playing in the new graphics mode.";
            soundenviron["Night Event"] = "Open - Play this sound at a random range and variance in the background when it is currently nighttime in the game.";
            soundenviron["HD Night Event"] = "Open - Play this sound at a random range and variance in the background when it is currently nighttime in the game while playing in the new graphics mode.";
            soundenviron["Event Delay"] = "Number - Controls the baseline number of frames to wait before playing the Day Event or Night Event sound, depending on the time of day. This only applies when the game is being played in SD mode. This value is used in the following calculation to get a random time to play the next event sound: [Event Delay] - [Event Delay] / 3 + RANDOM(0, ([Event Delay] / 3 * 2 + 1)).";
            soundenviron["HD Event Delay"] = "Number - Controls the baseline number of frames to wait before playing the Day Event or Night Event sound, depending on the time of day. This only applies when the game is being played in the new graphics mode. This value is used in the following calculation to get a random time to play the next event sound: [Event Delay] - [Event Delay] / 3 + RANDOM(0, ([Event Delay] / 3 * 2 + 1)).";
            soundenviron["Indoors"] = "Boolean - If equals 1 then, if the current sound being played in the area level with this Sound Environment is \"event_thunder_1\", then the sound will be obstructed. If equals 0, then ignore this.";
            soundenviron["Material 1"] = "Number - Controls the material of the Sound Environment, which affects which footstep sounds are played. Uses a code to define a specific material.";
            soundenviron["HD Material 1"] = "Number - Controls the material of the Sound Environment, which affects which footstep sounds are played. Uses a code to define a specific material. This only applies when the game is being played in the new graphics mode. See Material 1 for the code descriptions.";
            soundenviron["SFX EAX Environ"] = "Number - Determines an environment preset for default sound reverberation settings.";
            soundenviron["SFX EAX Room Vol"] = "Number - Room effect level at mid frequencies.";
            soundenviron["SFX EAX Room HF"] = "Number - Relative room effect level at high frequencies.";
            soundenviron["SFX EAX Decay Time"] = "Number - Reverberation decay time at mid frequencies.";
            soundenviron["SFX EAX Decay HF"] = "Number - High-frequency to mid-frequency decay time ratio.";
            soundenviron["SFX EAX Reflect"] = "Number - Early reflections level relative to room effect.";
            soundenviron["SFX EAX Reflect Delay"] = "Number - Initial reflection delay time.";
            soundenviron["SFX EAX Reverb"] = "Number - Late reverberation level relative to room effect.";
            soundenviron["SFX EAX Rev Delay"] = "Number - Late reverberation delay time relative to initial reflection.";
            soundenviron["VOX EAX Environ"] = "Number - Determines an environment preset for default sound reverberation settings.";
            soundenviron["VOX EAX Room Vol"] = "Number - Room effect level at mid frequencies.";
            soundenviron["VOX EAX Room HF"] = "Number - Relative room effect level at high frequencies.";
            soundenviron["VOX EAX Decay Time"] = "Number - Reverberation decay time at mid frequencies.";
            soundenviron["VOX EAX Decay HF"] = "Number - High-frequency to mid-frequency decay time ratio.";
            soundenviron["VOX EAX Reflect"] = "Number - Early reflections level relative to room effect.";
            soundenviron["VOX EAX Reflect Delay"] = "Number - Initial reflection delay time.";
            soundenviron["VOX EAX Reverb"] = "Number - Late reverberation level relative to room effect.";
            soundenviron["VOX EAX Rev Delay"] = "Number - Late reverberation delay time relative to initial reflection.";
            soundenviron["InheritEnvironment"] = "Boolean - If equals 1, then this sound environment will inherit certain values from the existing environment and overwrite other values with its own.";

            states["state"] = "Open - Defines the unique name ID for the state.";
            states["group"] = "Number - Assigns the state to a group ID value. This means that only 1 state with that group ID can be active at any time on a unit. If this value is empty, then ignore this.";
            states["remhit"] = "Boolean - If equals 1, then this state will be removed when the unit is hit. If equals 0, then ignore this.";
            states["nosend"] = "Boolean - If equals 1, then this state change will not be sent to the client. If equals 0, then ignore this.";
            states["transform"] = "Boolean - If equals 1, then this state will be flagged to change the unit's appearance and reset its animations when it is applied. If equals 0, then ignore this.";
            states["aura"] = "Boolean - If equals 1, then this state will be treated as an aura. If equals 0, then ignore this.";
            states["curable"] = "Boolean - If equals 1, then this state can be cured (This can be checked by NPC healing or the Paladin Cleansing skill). If equals 0, then ignore this.";
            states["curse"] = "Boolean - If equals 1, then this state will be flagged as a curse. If equals 0, then ignore this.";
            states["active"] = "Boolean - If equals 1, then the state will be classified as an active state which enables the cltactivefunc and srvactivefunc fields. If equals 0, then ignore this.";
            states["restrict"] = "Boolean - If equals 1, then this state will restrict the usage of certain skills (This connects with the restrict field from the Skills.txt file). If equals 0, then ignore this.";
            states["disguise"] = "Boolean - If equals 1, then this state will be flagged as a disguise, meaning that the unit's appearance is changed, which can affect how the animations are treated when being used. If equals 0, then ignored this.";
            states["attblue"] = "Boolean - If equals 1, then the state will make the related Attack Rating value in the character screen be colored blue. If equals 0, then ignore this.";
            states["damblue"] = "Boolean - If equals 1, then the state will make related Damage value in the character screen be colored blue. If equals 0, then ignore this.";
            states["armblue"] = "Boolean - If equals 1, then the state will make Defense value (Armor) in the character screen be colored blue. If equals 0, then ignore this.";
            states["rfblue"] = "Boolean - If equals 1, then the state will make Fire Resistance value in the character screen be colored blue. If equals 0, then ignore this.";
            states["rlblue"] = "Boolean - If equals 1, then the state will make Lightning Resistance value in the character screen be colored blue. If equals 0, then ignore this.";
            states["rcblue"] = "Boolean - If equals 1, then the state will make Cold Resistance value in the character screen be colored blue. If equals 0, then ignore this.";
            states["stambarblue"] = "Boolean - If equals 1, then the state will make the Stamina Bar UI in the HUD be colored blue. If equals 0, then ignore this.";
            states["rpblue"] = "Boolean - If equals 1, then the state will make Poison Resistance value in the character screen be colored blue. If equals 0, then ignore this.";
            states["attred"] = "Boolean - If equals 1, then the state will make the related Attack Rating value in the character screen be colored red. If equals 0, then ignore this.";
            states["damred"] = "Boolean - If equals 1, then the state will make related Damage value in the character screen be colored red. If equals 0, then ignore this.";
            states["armred"] = "Boolean - If equals 1, then the state will make Defense value (Armor) in the character screen be colored red. If equals 0, then ignore this.";
            states["rfred"] = "Boolean - If equals 1, then the state will make Fire Resistance value in the character screen be colored red. If equals 0, then ignore this.";
            states["rlred"] = "Boolean - If equals 1, then the state will make Lightning Resistance value in the character screen be colored red. If equals 0, then ignore this.";
            states["rcred"] = "Boolean - If equals 1, then the state will make Cold Resistance value in the character screen be colored red. If equals 0, then ignore this.";
            states["rpred"] = "Boolean - If equals 1, then the state will make Poison Resistance value in the character screen be colored red. If equals 0, then ignore this.";
            states["exp"] = "Boolean - If equals 1, then a unit with this state will give exp when killed or will gain exp when killing another unit. If equals 0, then ignore this.";
            states["plrstaydeath"] = "Boolean - If equals 1, then the state will persist on the player after that player is killed. If equals 0, then ignore this. state stays after death.";
            states["monstaydeath"] = "Boolean - If equals 1, then the state will persist on the monster (non-boss) after that monster is killed. If equals 0, then ignore this.";
            states["bossstaydeath"] = "Boolean - If equals 1, then the state will persist on the boss after that boss is killed. If equals 0, then ignore this.";
            states["hide"] = "Boolean - If equals 1, then the state will hide the unit when dead (corpse and death animations will not be drawn). If equals 0, then ignore this.";
            states["hidedead"] = "Boolean - If equals 1, then the state will be used to destroy units with invisible corpses. If equals 0, then ignore this.";
            states["shatter"] = "Boolean - If equals 1, then the state causes ice shatter missiles to create when the unit dies. If equals 0, then ignore this.";
            states["udead"] = "Boolean - If equals 1, then the state flags the unit as a used dead corpse and the unit cannot be targeted for corpse skills. If equals 0, then ignore this.";
            states["life"] = "Boolean - If equals 1, then this state will cancel out the monster's normal life regeneration. If equals 0, then ignore this.";
            states["green"] = "Boolean - If equals 1, then the state overrides the color changes the unit and the unit will be colored green. If equals 0, then ignore this.";
            states["pgsv"] = "Boolean - If equals 1, then the state is flagged as part of a progressive skill which relates to charge-up skill functionalities. If equals 0, then ignore this.";
            states["nooverlays"] = "Boolean - If equals 1, then the standard way for States to add overlays will be disabled. If equals 0, then ignore this.";
            states["noclear"] = "Boolean - If equals 1, then when this state is applied on the unit, it will not clear stats that have this state from the state's previous application. If equals 0, then ignore this.";
            states["bossinv"] = "Boolean - If equals 1, then the unit with this state will use the state's source unit's (in this case, the unit's boss) inventory for generating the unit's equipped item graphics. If equals 0, then ignore this.";
            states["meleeonly"] = "Boolean - If equals 1, then the state will make the all the unit's attack become melee attacks. If equals 0, then ignore this.";
            states["notondead"] = "Boolean - If equals 1, then the state will not play its On function (function that happens when the state is applied) if the unit is dead. If equals 0, then ignore this.";
            states["overlay1"] = "Open - Controls which overlay to use for normally displaying the state (Uses the overlay field from Overlay.txt). The usage depends on the specific state defined and/or the function using the state. Typically, States use \"overlay1\" for the Front overlay and \"overlay2\" for the Back overlay. Other cases can have States use each overlay field as the Front Start, Front End, Back Start, and Back End, respectively.";
            states["pgsvoverlay"] = "Open - Controls which overlay to use when the state has progressive charges on the unit, such as for the charge-up stat when using Assassin Martial Arts charge-up skills (Uses the overlay field from Overlay.txt).";
            states["castoverlay"] = "Open - Controls which overlay to use when the state is initially applied on the unit (Uses the overlay field from Overlay.txt).";
            states["removerlay"] = "Open - Controls which overlay to use when the state is removed from the unit (Uses the overlay field from Overlay.txt).";
            states["stat"] = "Open - Controls the stat associated with the stat. This is also used when determining how to add the progressive overlay (Uses the Stat field from ItemStatCost.txt).";
            states["setfunc"] = "Number - Controls the client side set functions for when the state is initially applied on the unit.";
            states["remfunc"] = "Number - Controls the client side remove functions for when the state is removed from the unit.";
            states["missile"] = "Open - Used as a possible parameter for the setfunc field (Uses the Missile field from Missiles.txt).";
            states["skill"] = "Open - Used as a possible parameter for the setfunc field (Uses the Skill field from Skills.txt).";
            states["itemtype"] = "Open - Defines a potential ItemType that can be affected by the state's color change.";
            states["itemtrans"] = "Open - Controls the color change of the item when the unit has this state. Referenced from the Code column in Colors.txt.";
            states["colorpri"] = "Number - Defines the priority of the state's color change, when compared to other current sates on the unit. The current state that has the highest color priority on the unit will be used and other state colors will be ignored. If multiple current States share the same color priority value, then the game will choose the state with the lower ID value (based on where in the list of States in the data file that the state is defined).";
            states["colorshift"] = "Number - Controls which index of the color shift palette to use.";
            states["light-r"] = "Number - Controls the state's change of the red color value of the Light radius (Uses a value from 0 to 255).";
            states["light-g"] = "Number - Controls the state's change of the green color value of the Light radius (Uses a value from 0 to 255).";
            states["light-b"] = "Number - Controls the state's change of the blue color value of the Light radius (Uses a value from 0 to 255).";
            states["onsound"] = "Open - Plays a sound when the state is initially applied to the unit. Links to a Sound from Sounds.txt.";
            states["offsound"] = "Open - Plays a sound when the state is removed from the unit. Links to a Sound from Sounds.txt.";
            states["gfxtype"] = "Number - Controls the how to handle the unit graphics transformation based on the unit type (This relies on the disguise field being enabled). If equals 1, then use this on a monster type unit. If equals 2, then use this on a player type unit. Otherwise, ignore this.";
            states["gfxclass"] = "Number - Control's the unit class used for handling the unit graphics transformation. This field relies on what unit type was used in the gfxtype field. If gfxtype equals 1 for monster type units, then this field will rely on the hcIDx field from MonStats.txt. If \"gfxtype\" equals 2, then this field will use the character class numeric ID.";
            states["cltevent"] = "Number - Controls the event to check on the client side to determine when to use the function defined in the clteventfunc field (Uses an event defined in the Events.txt file).";
            states["clteventfunc"] = "Number - Controls the client Unit event function that is called when the event is determined in the cltevent field (Uses the functions defined in the Events.txt file).";
            states["cltactivefunc"] = "Number - Controls the Client Do function that is called every frame while the state is active (the cltdofunc field in Skills.txt). This relies on the \"active\" field being enabled.";
            states["srvactivefunc"] = "Number - Controls the Server Do function that is called every frame while the state is active (the srvdofunc field in Skills.txt). This relies on the \"active\" field being enabled.";
            states["canstack"] = "Broken - Boolean - If equals 1, then this state can stack with duplicate forms of itself (This is only usable with the \"poison\" state). If equals 0, then ignore this.";
            states["sunderfull"] = "Broken - Boolean - If equals 1, then this state will reapply any negative resistance stats at full potential when calculating pierce immunity if the immunity was broken. If equals 0, then reapply at the normal reduced efficiency (currently 1/5).";

            superuniques["Superunique"] = "Open - Defines the unique name ID for the Super Unique monster.";
            superuniques["Name"] = "Open - Uses a string for the Super Unique monster's name.";
            superuniques["Class"] = "Open - Defines the baseline monster type for the Super Unique monster, which this monster will use for default values. This uses the ID field from MonStats.txt.";
            superuniques["hcIDx"] = "Number - Defines the unique numeric ID for the Super Unique monster. The existing IDs are hardcoded for specific scripts with the specified Super Unique monsters.";
            superuniques["MonSound"] = "Open - Defines what set of sounds to use for the Super Unique monster. Uses the ID field from MonSounds.txt. If this field is empty, then the Super Unique monster will default to using the monster class sounds.";
            superuniques["MinGrp"] = "Number - Controls the min amount of Minion monsters that will spawn with the Super Unique monster.";
            superuniques["MaxGrp"] = "Number - Controls the max amount of Minion monsters that will spawn with the Super Unique monster. This value must be equal to or higher than MinGrp. If this value is greater than it, then a random number will be chosen between the MinGrp and MaxGrp values.";
            superuniques["AutoPos"] = "Boolean - If equals 1, then the Super Unique monster will randomly spawn within a radius of its designated position. If equals 0, then the Super Unique monster will spawn at exact coordinates of its designated position.";
            superuniques["Stacks"] = "Boolean - If equals 1, then this Super Unique monster can spawn more than once in the same game. If equals 0, then this Super Unique monster can only spawn once in the same game.";
            superuniques["Replaceable"] = "Boolean - If equals 1, then the room where the Super Unique monster spawns in can be replaced during the creation of a level preset. If equals 0, then the room cannot be replaced and will remain static.";
            superuniques["Utrans & Utrans(N) & UTrans(H)"] = "Number - Modifies the color transform for the unique monster respectively in Normal, Nightmare, or Hell difficulty. If this value is greater than or equal to 30, then the value will default to 2, which is the monster's default color palette shift. If the value is 0 or is empty, then a random value will be chosen.";
            superuniques["TC & TC(N) & TC(H)"] = "Open - Controls the Treasure Class to use when the Super Unique monster is killed respectively in Normal, Nightmare, or Hell difficulty. This is linked to the corresponding Treasure Class.";

            treasureclassex["Treasure Class"] = "Open - Defines the unique Treasure Class ID, that is referenced in other files.";
            treasureclassex["group"] = "Number - Assigns the Treasure Class to a group ID value, which will connect this Treasure Class with other Treasure Classes, as a potential Treasure Class to use for an itemdrop. When determining which Treasure Class to use for an item drop, the game will iterate through all Treasure Classes that share the same group. This field works with the level field to determine an ideal Treasure Class to use for the monster drop. Treasure Classes that share the same group should be in contiguous order.";
            treasureclassex["level"] = "Number - Defines the level of a Treasure Class. Monsters who have a Treasure Class will pick the Treasure Class that a level value that is less than or equal to the monster's level. This is ignored for Boss monsters.";
            treasureclassex["Picks"] = "Number - Controls how to handle the calculations for item drops. If this value is positive, then this value will control how many item drop chances will be rolled for the Treasure Class using the \"Prob#\" fields as probability values. If this value is negative, then this value functions as the total guaranteed quantity of item drops from the Treasure Class, and each \"Prob#\" field now defines the quantity of items generated from its related \"Item#\" field. If this field is empty, then default to a value of 1.";
            treasureclassex["Unique"] = "Number - Modifies the item ratio drop for a Unique Quality item. A higher value means a better chance of being chosen. (ItemRatio.txt for an explanation for how the Item Quality is chosen).";
            treasureclassex["Set"] = "Number - Modifies the item ratio drop for a Set Quality item. A higher value means a better chance of being chosen. (ItemRatio.txt for an explanation for how the Item Quality is chosen).";
            treasureclassex["Rare"] = "Number - Modifies the item ratio drop for a Rare Quality item. A higher value means a better chance of being chosen. (ItemRatio.txt for an explanation for how the Item Quality is chosen).";
            treasureclassex["Magic"] = "Number - Modifies the item ratio drop for a Magic Quality item. A higher value means a better chance of being chosen. (ItemRatio.txt for an explanation for how the Item Quality is chosen).";
            treasureclassex["NoDrop"] = "Number - Controls the probability of no item dropping by the Treasure Class. The higher this value, then the more likely no item will drop from the monster. This can be automatically be affected by the number of players currently in the game.";
            treasureclassex["Item1"] = "Open - Defines a potential ItemType or other Treasure Class that can drop from this Treasure Class. Linking another Treasure Class in this field means that there is a chance to use that Treasure Class group of items which the game will then calculate a selection from that Treasure Class, and so on.";
            treasureclassex["Prob1"] = "Number - The individual probability for each related Item1 drop. The higher this value, then the more likely the Item1 field will be chosen. The chance a drop is picked is calculated by summing all Prob1 field values and the NoDrop value for a total denominator value, and then having each Prob1 value and the NoDrop value rolling their chance out of the total denominator value for a drop.";

            uniqueappellation["Name"] = "Open - A string key, which is used as a potential selection for generating a unique monster's name.";

            uniqueitems["index"] = "Open - Points to a string key value to use as the Unique item's name.";
            uniqueitems["version"] = "Number - Defines which game version to create this item (0 = Classic mode | 100 = Expansion mode).";
            uniqueitems["enabled"] = "Boolean - If equals 1, then this item can be created and dropped. If equals 0, then this item cannot be dropped.";
            uniqueitems["ladder"] = "Boolean - If equals 1, then this item can only be created and dropped in online Battle.net Ladder games. If equals 0, then this item can be created and dropped in any game mode.";
            uniqueitems["rarity"] = "Number - Modifies the chances that this Unique item will spawn compared to the other Unique items. This value acts as a numerator and a denominator. Each \"rarity\" value gets summed together to give a total denominator, used for the random roll for the item. For example, if there are 3 possible Unique items, and their \"rarity\" values are 3, 5, 7, then their chances to be chosen are 3/15, 5/15, and 7/15 respectively. (The minimum \"rarity\" value equals 1) (Only works for games in Expansion mode).";
            uniqueitems["nolimit"] = "Boolean - Requires the quest field from AMW.txt to be enabled. If equals 1, then this item can be created and will automatically be identified. If equals 0, then ignore this.";
            uniqueitems["lvl"] = "Number - The item level for the item, which controls what object or monster needs to be in order to drop this item.";
            uniqueitems["code"] = "Open - Defines the baseline item code to use for this Unique item (must match the code field from AMW.txt).";
            uniqueitems["carry1"] = "Boolean - If equals 1, then players can only carry one of these items in their inventory. If equals 0, then ignore this.";
            uniqueitems["cost mult"] = "Number - Multiplicative modifier for the Unique item's buy, sell, and repair costs.";
            uniqueitems["cost add"] = "Number - Flat integer modification to the Unique item's buy, sell, and repair costs. This is added after the cost mult has modified the costs.";
            uniqueitems["chrtransform"] = "Open - Controls the color change of the item when equipped on a character or dropped on the ground. If empty, then the item will have the default item color. Referenced from the Code column in Colors.txt.";
            uniqueitems["invtransform"] = "Open - Controls the color change of the item in the inventory UI. If empty, then the item will have the default item color. Referenced from the Code column in Colors.txt.";
            uniqueitems["invfile"] = "Open - An override for the invfile field from AMW.txt. By default, the Unique item will use what was defined by the baseline item from the \"item\" field.";
            uniqueitems["flippyfile"] = "Open - An override for the \"flippyfile\" field from the weapon.txt, Armor.txt, or Misc.txt files. By default, the Unique item will use what was defined by the baseline item from the \"item\" field.";
            uniqueitems["dropsound"] = "Open - An override for the \"dropsound\" field from the weapon.txt, Armor.txt, or Misc.txt files. By default, the Unique item will use what was defined by the baseline item from the \"item\" field.";
            uniqueitems["dropsfxframe"] = "Number - An override for the \"dropsfxframe\" field from the weapon.txt, Armor.txt, or Misc.txt files. By default, the Unique item will use what was defined by the baseline item from the \"item\" field.";
            uniqueitems["usesound"] = "Open - An override for the \"usesound\" field from the weapon.txt, Armor.txt, or Misc.txt files. By default, the Unique item will use what was defined by the baseline item from the \"item\" field.";
            uniqueitems["prop1"] = "Open - Controls the item properties for the Unique item (Uses the Code field from Properties.txt).";
            uniqueitems["par1"] = "Number - The stat's \"parameter\" value associated with the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";
            uniqueitems["min1"] = "Number - The stat's \"min\" value to assign to the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";
            uniqueitems["max1"] = "Number - The stat's \"max\" value to assign to the related property (prop#). Usage depends on the (Function ID field from Properties.txt).";
            uniqueitems["diablocloneweight"] = "Number - The amount of weight added to the diablo clone progress when this item is sold. When offline, selling this item will instead immediately spawn diablo clone.";

            uniqueprefix["Name"] = "Open - A string key, which is used as a potential selection for generating a unique monster's name.";

            uniquesuffix["Name"] = "Open - A string key, which is used as a potential selection for generating a unique monster's name.";

            wanderingmon["class"] = "Open - Uses a monster \"ID\" defined from the monstats.txt file. Monsters defined here are added to a list which is used to randomly pick a monster to spawn in an area level.";

            result["ActInfo.txt"] = actinfo;
            result["Armor.txt"] = armor;
            result["Misc.txt"] = misc;
            result["Weapons.txt"] = weapons;
            result["AutoMagic.txt"] = automagic;
            result["AutomMap.txt"] = automap;
            result["Books.txt"] = books;
            result["Charstats.txt"] = charstats;
            result["Cubemain.txt"] = cubemain;
            result["DifficultyLevels.txt"] = difficultylevels;
            result["Experience.txt"] = experience;
            result["Gamble.txt"] = gamble;
            result["Gems.txt"] = gems;
            result["Hireling.txt"] = hireling;
            result["HirelingDesc.txt"] = hirelingdesc;
            result["Inventory.txt"] = inventory;
            result["ItemRatio.txt"] = itemratio;
            result["ItemStatCost.txt"] = itemstatcost;
            result["ItemTypes.txt"] = itemtypes;
            result["LevelGroups.txt"] = levelgroups;
            result["Levels.txt"] = levels;
            result["LvlMaze.txt"] = lvlmaze;
            result["LvlPrest.txt"] = lvlprest;
            result["LvlSub.txt"] = lvlsub;
            result["LvlTypes.txt"] = lvltypes;
            result["LvlWarp.txt"] = lvlwarp;
            result["MagicPrefix.txt"] = magicprefix;
            result["MagicSuffix.txt"] = magicsuffix;
            result["Missiles.txt"] = missiles;
            result["MonEquip.txt"] = monequip;
            result["MonLvl.txt"] = monlvl;
            result["MonPreset.txt"] = monpreset;
            result["MonProp.txt"] = monprop;
            result["MonSeq.txt"] = monseq;
            result["MonStats.txt"] = monstats;
            result["MonStats2.txt"] = monstats2;
            result["MonType.txt"] = montype;
            result["MonUMod.txt"] = monumod;
            result["MonSounds.txt"] = monsounds;
            result["NPC.txt"] = npc;
            result["Objects.txt"] = objects;
            result["ObjGroup.txt"] = objgroup;
            result["ObjPreset.txt"] = objpreset;
            result["Overlay.txt"] = overlay;
            result["PetType.txt"] = pettype;
            result["Properties.txt"] = properties;
            result["QualityItems.txt"] = qualityitems;
            result["RarePrefix.txt"] = rareprefix;
            result["RareSuffix.txt"] = raresuffix;
            result["Runes.txt"] = runes;
            result["SetItems.txt"] = setitems;
            result["Sets.txt"] = sets;
            result["Shrines.txt"] = shrines;
            result["Skills.txt"] = skills;
            result["SkillDesc.txt"] = skilldesc;
            result["Sounds.txt"] = sounds;
            result["SoundEnviron.txt"] = soundenviron;
            result["States.txt"] = states;
            result["SuperUniques.txt"] = superuniques;
            result["TreasureClassEx.txt"] = treasureclassex;
            result["UniqueAppellation.txt"] = uniqueappellation;
            result["UniqueItems.txt"] = uniqueitems;
            result["UniquePrefix.txt"] = uniqueprefix;
            result["UniqueSuffix.txt"] = uniquesuffix;
            result["WanderingMon.txt"] = wanderingmon;

            return result;
        }

        private static string? GetManualHeaderTooltip(string fileName, string columnName)
        {
            if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(columnName)) return null;
            if (!ManualHeaderTooltips.TryGetValue(fileName, out var byColumn)) return null;
            return byColumn.TryGetValue(columnName, out string? text) ? text : null;
        }

        private static readonly string[] HeaderTooltipMarkers = { "Boolean", "Open", "Number", "Calculation", "Comment", "Broken" };

        private const int HeaderTooltipPaddingH = 6;
        private const int HeaderTooltipPaddingV = 4;
        private const int HeaderTooltipContentWidth = 420;

        private void LayoutHeaderTooltipPanel()
        {
            if (_headerTooltipPanel == null || _headerTooltipLabelBold == null || _headerTooltipLabelNormal == null) return;

            Size boldPref = _headerTooltipLabelBold.GetPreferredSize(new Size(HeaderTooltipContentWidth, int.MaxValue));
            if (boldPref.Width <= 0 || boldPref.Height <= 0)
            {
                int lineHeight = _headerTooltipLabelBold.Font.Height + 2;
                boldPref = new Size(Math.Max(boldPref.Width, 60), Math.Max(boldPref.Height, lineHeight));
            }
            _headerTooltipLabelBold.Size = boldPref;
            _headerTooltipPanel.RowStyles[0].SizeType = SizeType.Absolute;
            _headerTooltipPanel.RowStyles[0].Height = 15;

            _headerTooltipPanel.PerformLayout();
            // Compute size explicitly so it's correct even when panel has never been visible
            int row0Height = _headerTooltipPanel.RowStyles[0].SizeType == SizeType.Absolute
                ? (int)_headerTooltipPanel.RowStyles[0].Height
                : _headerTooltipLabelBold.Visible ? _headerTooltipLabelBold.Height : 0;
            int row1Height = _headerTooltipLabelNormal.Height;
            int totalHeight = _headerTooltipPanel.Padding.Vertical + row0Height + row1Height;
            int totalWidth = _headerTooltipPanel.Padding.Horizontal + HeaderTooltipContentWidth;
            _headerTooltipPanel.Size = new Size(totalWidth, totalHeight);
        }

        private static void SetHeaderTooltipContent(Label labelBold, Label labelNormal, string tooltipText)
        {
            if (string.IsNullOrEmpty(tooltipText))
            {
                labelBold.Text = "";
                labelBold.Visible = false;
                labelNormal.Text = "";
                labelNormal.Height = 0;
                return;
            }
            string markerUsed = "";
            int markerLen = 0;
            foreach (string m in HeaderTooltipMarkers)
            {
                string prefix = m + " - ";
                if (tooltipText.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    markerUsed = m;
                    markerLen = prefix.Length;
                    break;
                }
            }
            if (markerLen == 0)
            {
                labelBold.Visible = false;
                labelNormal.Text = tooltipText;
                labelNormal.Width = HeaderTooltipContentWidth;
                labelNormal.Height = labelNormal.GetPreferredSize(new Size(HeaderTooltipContentWidth, int.MaxValue)).Height;
                return;
            }
            labelBold.Text = markerUsed;
            labelBold.Visible = true;
            labelNormal.Text = tooltipText.Substring(markerLen);
            labelNormal.Width = HeaderTooltipContentWidth;
            labelNormal.Height = labelNormal.GetPreferredSize(new Size(HeaderTooltipContentWidth, int.MaxValue)).Height;
        }

        private void DataGuideHeaderMouseMoveFromTabOrGrid(object? sender, MouseEventArgs e)
        {
            if (!_headerTooltipsFromDataGuide) return;
            DataGridView? dgv = null;
            int clientX, clientY;
            if (sender is DataGridView grid)
            {
                dgv = grid;
                clientX = e.X;
                clientY = e.Y;
            }
            else if (sender is TabPage tab && tab.Controls.Count > 0 && tab.Controls[0] is DataGridView tabGrid)
            {
                dgv = tabGrid;
                Point screen = Cursor.Position;
                Point client = dgv.PointToClient(screen);
                clientX = client.X;
                clientY = client.Y;
                if (!dgv.ClientRectangle.Contains(clientX, clientY))
                {
                    if (_lastHeaderTooltipShown != null && _lastHeaderTooltipControl != null)
                    {
                        _lastHeaderTooltipShown = null;
                        toolTip1.Hide(_lastHeaderTooltipControl);
                        _lastHeaderTooltipControl = null;
                    }
                    return;
                }
            }
            else
            {
                if (_lastHeaderTooltipShown != null && _lastHeaderTooltipControl != null)
                {
                    _lastHeaderTooltipShown = null;
                    toolTip1.Hide(_lastHeaderTooltipControl);
                    _lastHeaderTooltipControl = null;
                }
                return;
            }
            if (dgv == null) return;
            var hit = dgv.HitTest(clientX, clientY);
            if (hit.Type != DataGridViewHitTestType.ColumnHeader || hit.ColumnIndex < 0)
            {
                if (_lastHeaderTooltipShown != null && _lastHeaderTooltipControl != null)
                {
                    _lastHeaderTooltipShown = null;
                    toolTip1.Hide(_lastHeaderTooltipControl);
                    _lastHeaderTooltipControl = null;
                }
                return;
            }
            string? filePath = GetFilePathForGrid(dgv);
            if (string.IsNullOrEmpty(filePath))
            {
                if (_lastHeaderTooltipShown != null && _lastHeaderTooltipControl != null)
                {
                    _lastHeaderTooltipShown = null;
                    toolTip1.Hide(_lastHeaderTooltipControl);
                    _lastHeaderTooltipControl = null;
                }
                return;
            }
            string fileName = Path.GetFileName(filePath);
            var col = dgv.Columns[hit.ColumnIndex];
            string columnName = col.HeaderText?.Trim() ?? col.Name ?? "";
            if (string.IsNullOrEmpty(columnName) || columnName == OriginalIndexColumnName)
            {
                if (_lastHeaderTooltipShown != null && _lastHeaderTooltipControl != null)
                {
                    _lastHeaderTooltipShown = null;
                    toolTip1.Hide(_lastHeaderTooltipControl);
                    _lastHeaderTooltipControl = null;
                }
                return;
            }
            string? description = GetManualHeaderTooltip(fileName, columnName);
            string tooltipText = !string.IsNullOrEmpty(description) ? description : "no info found";
            if (_lastHeaderTooltipShown is (int c, string t) && c == hit.ColumnIndex && string.Equals(t, tooltipText, StringComparison.Ordinal))
                return;
            _lastHeaderTooltipShown = (hit.ColumnIndex, tooltipText);
            _lastHeaderTooltipControl = dgv;
            toolTip1.SetToolTip(dgv, tooltipText);
            toolTip1.Show(tooltipText, dgv, clientX, clientY + 22, 10000);
        }

        private void LoadGridColors()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key == null) return;
                    if (key.GetValue("GridHeaderColorArgb") is int headerArgb)
                        _gridHeaderColor = Color.FromArgb(headerArgb);
                    if (key.GetValue("GridFrozenColorArgb") is int frozenArgb)
                        _gridFrozenColor = Color.FromArgb(frozenArgb);
                    if (key.GetValue("GridEvenRowColorArgb") is int evenArgb)
                        _gridEvenRowColor = Color.FromArgb(evenArgb);
                    if (key.GetValue("GridOddRowColorArgb") is int oddArgb)
                        _gridOddRowColor = Color.FromArgb(oddArgb);
                    if (key.GetValue("GridSearchHighlightColorArgb") is int searchArgb)
                        _gridSearchHighlightColor = Color.FromArgb(searchArgb);
                    if (key.GetValue("GridExpansionRowColorArgb") is int expansionArgb)
                        _gridExpansionRowColor = Color.FromArgb(expansionArgb);
                }
            }
            catch { }
        }

        private void SaveGridColors()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    key.SetValue("GridHeaderColorArgb", _gridHeaderColor.ToArgb());
                    key.SetValue("GridFrozenColorArgb", _gridFrozenColor.ToArgb());
                    key.SetValue("GridEvenRowColorArgb", _gridEvenRowColor.ToArgb());
                    key.SetValue("GridOddRowColorArgb", _gridOddRowColor.ToArgb());
                    key.SetValue("GridSearchHighlightColorArgb", _gridSearchHighlightColor.ToArgb());
                    key.SetValue("GridExpansionRowColorArgb", _gridExpansionRowColor.ToArgb());
                }
            }
            catch { }
        }

        private void LoadCustomPaletteColors()
        {
            try
            {
                _customPaletteColors.Clear();
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key?.GetValue("CustomPaletteColorsArgb") is string s && !string.IsNullOrWhiteSpace(s))
                    {
                        foreach (string part in s.Split(','))
                            if (int.TryParse(part.Trim(), out int argb))
                                _customPaletteColors.Add(Color.FromArgb(argb));
                    }
                }
            }
            catch { }
        }

        private void SaveCustomPaletteColors()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    string value = string.Join(",", _customPaletteColors.Select(c => c.ToArgb()));
                    key.SetValue("CustomPaletteColorsArgb", value);
                }
            }
            catch { }
        }

        private void RemoveD2CompareWorkspaces()
        {
            try
            {
                var nodesToRemove = new List<TreeNode>();
                foreach (TreeNode node in treeView.Nodes)
                {
                    if (node.Text.StartsWith("D2Compare (", StringComparison.OrdinalIgnoreCase))
                        nodesToRemove.Add(node);
                }

                foreach (var node in nodesToRemove)
                    treeView.Nodes.Remove(node);
            }
            catch { }
        }

        private void LoadTextColors()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key == null) return;
                    if (key.GetValue("TextColorTabArgb") is int v) _textColorTab = Color.FromArgb(v);
                    if (key.GetValue("TextColorHeaderArgb") is int v2) _textColorHeader = Color.FromArgb(v2);
                    if (key.GetValue("TextColorCellArgb") is int v3) _textColorCell = Color.FromArgb(v3);
                    if (key.GetValue("TextColorCellFrozenArgb") is int v3f) _textColorCellFrozen = Color.FromArgb(v3f);
                    if (key.GetValue("TextColorWorkspaceArgb") is int v4) _textColorWorkspace = Color.FromArgb(v4);
                    if (key.GetValue("TextColorEditHistoryArgb") is int v5) _textColorEditHistory = Color.FromArgb(v5);
                    if (key.GetValue("TextColorWorkspaceEntryArgb") is int ve) _textColorWorkspaceEntry = Color.FromArgb(ve);
                    if (key.GetValue("TextColorWorkspaceFilesArgb") is int vf) _textColorWorkspaceFiles = Color.FromArgb(vf);
                    if (key.GetValue("TextColorMenuStandbyArgb") is int vs) _textColorMenuStandby = Color.FromArgb(vs);
                    if (key.GetValue("TextColorMenuActiveArgb") is int va) _textColorMenuActive = Color.FromArgb(va);
                    if (key.GetValue("TextColorButtonsArgb") is int vb) _textColorButtons = Color.FromArgb(vb);
                    if (key.GetValue("TextColorTabDarkArgb") is int d1) _textColorTabDark = Color.FromArgb(d1);
                    if (key.GetValue("TextColorHeaderDarkArgb") is int d2) _textColorHeaderDark = Color.FromArgb(d2);
                    if (key.GetValue("TextColorCellDarkArgb") is int d3) _textColorCellDark = Color.FromArgb(d3);
                    if (key.GetValue("TextColorCellFrozenDarkArgb") is int d3f) _textColorCellFrozenDark = Color.FromArgb(d3f);
                    if (key.GetValue("TextColorWorkspaceDarkArgb") is int d4) _textColorWorkspaceDark = Color.FromArgb(d4);
                    if (key.GetValue("TextColorEditHistoryDarkArgb") is int d5) _textColorEditHistoryDark = Color.FromArgb(d5);
                    if (key.GetValue("TextColorWorkspaceEntryDarkArgb") is int de) _textColorWorkspaceEntryDark = Color.FromArgb(de);
                    if (key.GetValue("TextColorWorkspaceFilesDarkArgb") is int df) _textColorWorkspaceFilesDark = Color.FromArgb(df);
                    if (key.GetValue("TextColorMenuStandbyDarkArgb") is int ds) _textColorMenuStandbyDark = Color.FromArgb(ds);
                    if (key.GetValue("TextColorMenuActiveDarkArgb") is int da) _textColorMenuActiveDark = Color.FromArgb(da);
                    if (key.GetValue("TextColorButtonsDarkArgb") is int db) _textColorButtonsDark = Color.FromArgb(db);
                }
            }
            catch { }
        }

        private void SaveTextColors()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    key.SetValue("TextColorTabArgb", _textColorTab.ToArgb());
                    key.SetValue("TextColorHeaderArgb", _textColorHeader.ToArgb());
                    key.SetValue("TextColorCellArgb", _textColorCell.ToArgb());
                    key.SetValue("TextColorCellFrozenArgb", _textColorCellFrozen.ToArgb());
                    key.SetValue("TextColorWorkspaceArgb", _textColorWorkspace.ToArgb());
                    key.SetValue("TextColorEditHistoryArgb", _textColorEditHistory.ToArgb());
                    key.SetValue("TextColorWorkspaceEntryArgb", _textColorWorkspaceEntry.ToArgb());
                    key.SetValue("TextColorWorkspaceFilesArgb", _textColorWorkspaceFiles.ToArgb());
                    key.SetValue("TextColorMenuStandbyArgb", _textColorMenuStandby.ToArgb());
                    key.SetValue("TextColorMenuActiveArgb", _textColorMenuActive.ToArgb());
                    key.SetValue("TextColorButtonsArgb", _textColorButtons.ToArgb());
                    key.SetValue("TextColorTabDarkArgb", _textColorTabDark.ToArgb());
                    key.SetValue("TextColorHeaderDarkArgb", _textColorHeaderDark.ToArgb());
                    key.SetValue("TextColorCellDarkArgb", _textColorCellDark.ToArgb());
                    key.SetValue("TextColorCellFrozenDarkArgb", _textColorCellFrozenDark.ToArgb());
                    key.SetValue("TextColorWorkspaceDarkArgb", _textColorWorkspaceDark.ToArgb());
                    key.SetValue("TextColorEditHistoryDarkArgb", _textColorEditHistoryDark.ToArgb());
                    key.SetValue("TextColorWorkspaceEntryDarkArgb", _textColorWorkspaceEntryDark.ToArgb());
                    key.SetValue("TextColorWorkspaceFilesDarkArgb", _textColorWorkspaceFilesDark.ToArgb());
                    key.SetValue("TextColorMenuStandbyDarkArgb", _textColorMenuStandbyDark.ToArgb());
                    key.SetValue("TextColorMenuActiveDarkArgb", _textColorMenuActiveDark.ToArgb());
                    key.SetValue("TextColorButtonsDarkArgb", _textColorButtonsDark.ToArgb());
                }
            }
            catch { }
        }

        private void TextColorControlsMenuItem_Click(object? sender, EventArgs e)
        {
            using (var form = new TextColorControlsForm(_textColorTab, _textColorHeader, _textColorCell, _textColorCellFrozen, _textColorWorkspaceEntry, _textColorWorkspaceFiles, _textColorEditHistory, _textColorMenuStandby, _textColorMenuActive, _textColorButtons, _textColorTabDark, _textColorHeaderDark, _textColorCellDark, _textColorCellFrozenDark, _textColorWorkspaceEntryDark, _textColorWorkspaceFilesDark, _textColorEditHistoryDark, _textColorMenuStandbyDark, _textColorMenuActiveDark, _textColorButtonsDark))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;
                _textColorTab = form.TabTextColor;
                _textColorHeader = form.HeaderTextColor;
                _textColorCell = form.CellTextColor;
                _textColorCellFrozen = form.CellFrozenTextColor;
                _textColorWorkspaceEntry = form.WorkspaceEntryTextColor;
                _textColorWorkspaceFiles = form.WorkspaceFilesTextColor;
                _textColorEditHistory = form.EditHistoryTextColor;
                _textColorMenuStandby = form.MenuStandbyTextColor;
                _textColorMenuActive = form.MenuActiveTextColor;
                _textColorButtons = form.ButtonTextColor;
                _textColorTabDark = form.TabTextColorDark;
                _textColorHeaderDark = form.HeaderTextColorDark;
                _textColorCellDark = form.CellTextColorDark;
                _textColorCellFrozenDark = form.CellFrozenTextColorDark;
                _textColorWorkspaceEntryDark = form.WorkspaceEntryTextColorDark;
                _textColorWorkspaceFilesDark = form.WorkspaceFilesTextColorDark;
                _textColorEditHistoryDark = form.EditHistoryTextColorDark;
                _textColorMenuStandbyDark = form.MenuStandbyTextColorDark;
                _textColorMenuActiveDark = form.MenuActiveTextColorDark;
                _textColorButtonsDark = form.ButtonTextColorDark;
                SaveTextColors();
                tabControl.Invalidate();
                ApplyTextColors();
                foreach (TabPage tab in tabControl.TabPages)
                {
                    var dgv = GetGridForTab(tab);
                    if (dgv != null)
                    {
                        ApplyGridColorsToGrid(dgv);
                        ApplyRowStyles(dgv);
                        if (dgv == dataGridView)
                            ApplyHeaderCellColors(dgv);
                        dgv.Refresh();
                    }
                }
            }
        }

        private void ColorControlsMenuItem_Click(object? sender, EventArgs e)
        {
            using (var form = new ColorControlsForm(_gridHeaderColor, _gridFrozenColor, _gridEvenRowColor, _gridOddRowColor, _gridSearchHighlightColor, _gridExpansionRowColor))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;
                _gridHeaderColor = form.HeaderColor;
                _gridFrozenColor = form.FrozenColor;
                _gridEvenRowColor = form.EvenRowColor;
                _gridOddRowColor = form.OddRowColor;
                _gridSearchHighlightColor = form.SearchHighlightColor;
                _gridExpansionRowColor = form.ExpansionRowColor;
                SaveGridColors();
                foreach (TabPage tab in tabControl.TabPages)
                {
                    var dgv = GetGridForTab(tab);
                    if (dgv != null)
                    {
                        ApplyGridColorsToGrid(dgv);
                        ApplyRowStyles(dgv);
                        dgv.Refresh();
                        if (dgv.ColumnHeadersVisible)
                            dgv.Invalidate(new Rectangle(0, 0, dgv.Width, dgv.ColumnHeadersHeight));
                        dgv.Invalidate();
                    }
                }
            }
        }

        private void SaveOpenFiles()
        {
            try
            {
                var paths = new List<string>();
                foreach (TabPage tab in tabControl.TabPages)
                {
                    if (tab.Tag is TabTag tag && !string.IsNullOrEmpty(tag.FilePath))
                        paths.Add(tag.FilePath);
                }
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    if (paths.Count > 0)
                    {
                        key.SetValue("OpenFiles", paths.ToArray(), RegistryValueKind.MultiString);
                        key.SetValue("OpenFilesSelectedIndex", tabControl.SelectedIndex);
                    }
                    else
                    {
                        key.DeleteValue("OpenFiles", false);
                        key.DeleteValue("OpenFilesSelectedIndex", false);
                    }
                }
            }
            catch { }
        }

        private void LoadOpenFiles()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key?.GetValue("OpenFiles") is not string[] paths || paths.Length == 0)
                        return;
                    int selectedIndex = (int)key.GetValue("OpenFilesSelectedIndex", 0);
                    LoadOpenFilesNext(0, paths, selectedIndex);
                }
            }
            catch { }
        }

        private void LoadOpenFilesNext(int index, string[] paths, int selectedIndexToRestore)
        {
            if (index >= paths.Length)
            {
                if (tabControl.TabPages.Count > 0 && selectedIndexToRestore >= 0)
                    tabControl.SelectedIndex = Math.Min(selectedIndexToRestore, tabControl.TabPages.Count - 1);
                DisableColumnSorting();
                return;
            }
            string path = paths[index].Trim();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                LoadOpenFilesNext(index + 1, paths, selectedIndexToRestore);
                return;
            }
            var (wsColor, wsFolder) = GetWorkspaceColorForFilePath(path);
            LoadFile(path, wsColor, wsFolder, onComplete: () => LoadOpenFilesNext(index + 1, paths, selectedIndexToRestore));
        }

        private void SaveExpandedWorkspaces()
        {
            try
            {
                var expanded = new List<string>();
                foreach (TreeNode node in treeView.Nodes)
                {
                    if (node.Parent == null && node.IsExpanded && node.Tag is WorkspaceTag tag && !string.IsNullOrEmpty(tag.FolderPath))
                        expanded.Add(NormalizeFolderPath(tag.FolderPath));
                }
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
                {
                    if (expanded.Count > 0)
                        key.SetValue("ExpandedWorkspaces", expanded.ToArray(), RegistryValueKind.MultiString);
                    else
                        key.DeleteValue("ExpandedWorkspaces", false);
                }
            }
            catch { }
        }

        private void LoadExpandedWorkspaces()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key?.GetValue("ExpandedWorkspaces") is not string[] paths || paths.Length == 0)
                        return;
                    var expandedSet = new HashSet<string>(paths.Select(NormalizeFolderPath), StringComparer.OrdinalIgnoreCase);
                    foreach (TreeNode node in treeView.Nodes)
                    {
                        if (node.Parent == null && node.Tag is WorkspaceTag tag && !string.IsNullOrEmpty(tag.FolderPath)
                            && expandedSet.Contains(NormalizeFolderPath(tag.FolderPath)))
                            node.Expand();
                    }
                }
            }
            catch { }
        }

        private void SetDoubleBuffered(Control control)
        {
            var propertyInfo = control.GetType().GetProperty("DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            if (propertyInfo != null)
                propertyInfo.SetValue(control, true);
        }

    }

    public class ColorPaletteForm : Form
    {
        public Color ChosenColor { get; private set; }
        private const int SwatchSize = 32;
        private const int PanelPadding = 8;
        private readonly FlowLayoutPanel _swatchPanel;
        private readonly List<Color>? _customColors;
        private readonly Button _otherBtn;
        public ColorPaletteForm(Color currentColor, List<Color> customColors)
        {
            ChosenColor = currentColor;
            _customColors = customColors;
            Text = "Choose workspace color";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            int initialCount = D2Horadrim.DefaultWorkspaceColors.Length + (customColors?.Count ?? 0) + 1;
            int rowCount = 1 + (int)Math.Ceiling((initialCount + 2) / 6.0); // +2 for OK/Cancel
            Size = new Size(260, Math.Min(420, PanelPadding * 2 + rowCount * (SwatchSize + 4) + 40));
            _swatchPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(PanelPadding),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };
            void AddSwatch(Color c, bool isSelected)
            {
                var btn = new Button
                {
                    Size = new Size(SwatchSize, SwatchSize),
                    BackColor = c,
                    FlatStyle = FlatStyle.Flat,
                    Margin = new Padding(2)
                };
                btn.FlatAppearance.BorderSize = isSelected ? 3 : 1;
                btn.FlatAppearance.BorderColor = Color.DarkGray;
                btn.Click += (s, _) =>
                {
                    ChosenColor = (s as Button)!.BackColor;
                    DialogResult = DialogResult.OK;
                    Close();
                };
                _swatchPanel.Controls.Add(btn);
            }
            foreach (Color c in D2Horadrim.DefaultWorkspaceColors)
                AddSwatch(c, c.ToArgb() == currentColor.ToArgb());
            if (customColors != null)
            {
                foreach (Color c in customColors)
                    AddSwatch(c, c.ToArgb() == currentColor.ToArgb());
            }
            _otherBtn = new Button
            {
                Text = "Other...",
                Size = new Size(60, SwatchSize),
                Margin = new Padding(2)
            };
            _otherBtn.Click += (s, _) =>
            {
                using (var dlg = new ColorDialog { Color = ChosenColor, FullOpen = true })
                {
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        ChosenColor = dlg.Color;
                        if (_customColors != null && !_customColors.Any(c => c.ToArgb() == dlg.Color.ToArgb()))
                        {
                            _customColors.Add(dlg.Color);
                            var newBtn = new Button
                            {
                                Size = new Size(SwatchSize, SwatchSize),
                                BackColor = dlg.Color,
                                FlatStyle = FlatStyle.Flat,
                                Margin = new Padding(2)
                            };
                            newBtn.FlatAppearance.BorderSize = 3;
                            newBtn.FlatAppearance.BorderColor = Color.DarkGray;
                            newBtn.Click += (s2, _) =>
                            {
                                ChosenColor = (s2 as Button)!.BackColor;
                                DialogResult = DialogResult.OK;
                                Close();
                            };
                            int insertAt = _swatchPanel.Controls.IndexOf(_otherBtn);
                            _swatchPanel.Controls.Add(newBtn);
                            _swatchPanel.Controls.SetChildIndex(newBtn, insertAt);
                        }
                    }
                }
            };
            _swatchPanel.Controls.Add(_otherBtn);
            var okBtn = new Button { Text = "OK", Size = new Size(75, 28), Margin = new Padding(2) };
            okBtn.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
            var cancelBtn = new Button { Text = "Cancel", Size = new Size(75, 28), Margin = new Padding(2) };
            cancelBtn.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
            _swatchPanel.Controls.Add(okBtn);
            _swatchPanel.Controls.Add(cancelBtn);
            Controls.Add(_swatchPanel);
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }
    }

    public class ConfigurationForm : Form
    {
        public Font? SelectedFont { get; private set; }
        public bool CreateBackupOnSave { get; private set; }
        public int BackupFileCount { get; private set; }
        private Font? _currentFont;
        private Label? _fontLabel;
        private NumericUpDown? _backupCountNud;

        public ConfigurationForm(Font? currentFont, bool createBackupOnSave, int backupFileCount)
        {
            _currentFont = currentFont != null ? (Font)currentFont.Clone() : new Font(SystemFonts.DefaultFont.FontFamily, 9f);
            CreateBackupOnSave = createBackupOnSave;
            BackupFileCount = Math.Max(1, Math.Min(20, backupFileCount));
            SelectedFont = currentFont != null ? (Font)currentFont.Clone() : null;
            Text = "Configuration";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            Padding = new Padding(10);
            int tableHeight = 22 + 28 + 28 + 34 + 36;
            Size = new Size(240, tableHeight + 20 + 39);
            MinimumSize = new Size(240, Size.Height);

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(0),
                ColumnCount = 2,
                RowCount = 5,
                Height = tableHeight
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            var fontLinePanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                WrapContents = false
            };
            var fontDescLabel = new Label { Text = "Font:", AutoSize = true, Margin = new Padding(0, 2, 8, 0) };
            _fontLabel = new Label
            {
                Text = _currentFont.Name + ", " + _currentFont.SizeInPoints.ToString("F1") + " pt",
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 0),
                AutoEllipsis = true
            };
            fontLinePanel.Controls.Add(fontDescLabel);
            fontLinePanel.Controls.Add(_fontLabel);
            table.Controls.Add(fontLinePanel, 0, 0);
            table.SetColumnSpan(fontLinePanel, 2);

            var fontBtnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                WrapContents = false,
                Margin = new Padding(0, -3, 0, 0)
            };
            var fontChangeBtn = new Button { Text = "Change…", Size = new Size(62, 24), Margin = new Padding(0, 0, 8, 0) };
            fontChangeBtn.Click += (_, _) =>
            {
                using (var dlg = new FontDialog { Font = _currentFont ?? SystemFonts.DefaultFont })
                {
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        _currentFont?.Dispose();
                        _currentFont = (Font)dlg.Font.Clone();
                        if (_fontLabel != null)
                            _fontLabel.Text = _currentFont.Name + ", " + _currentFont.SizeInPoints.ToString("F1") + " pt";
                    }
                }
            };
            var fontDefaultBtn = new Button { Text = "Default", Size = new Size(62, 24), Margin = new Padding(0, 0, 8, 0) };
            fontDefaultBtn.Click += (_, _) =>
            {
                _currentFont?.Dispose();
                _currentFont = new Font(SystemFonts.DefaultFont.FontFamily, 9f);
                if (_fontLabel != null)
                    _fontLabel.Text = _currentFont.Name + ", " + _currentFont.SizeInPoints.ToString("F1") + " pt";
            };
            fontBtnPanel.Controls.Add(fontChangeBtn);
            fontBtnPanel.Controls.Add(fontDefaultBtn);
            table.Controls.Add(fontBtnPanel, 0, 1);
            table.SetColumnSpan(fontBtnPanel, 2);

            var backupCheckBox = new CheckBox
            {
                Text = "Backup on save",
                AutoSize = true,
                Checked = createBackupOnSave,
                Anchor = AnchorStyles.Left
            };
            table.Controls.Add(backupCheckBox, 0, 2);
            table.SetColumnSpan(backupCheckBox, 2);

            var backupCountLabel = new Label { Text = "Backups to keep:", AutoSize = true, Anchor = AnchorStyles.Left };
            _backupCountNud = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 20,
                Value = BackupFileCount,
                Width = 48,
                Anchor = AnchorStyles.Left,
                Enabled = backupCheckBox.Checked
            };
            backupCheckBox.CheckedChanged += (_, _) => _backupCountNud!.Enabled = backupCheckBox.Checked;
            table.Controls.Add(backupCountLabel, 0, 3);
            table.Controls.Add(_backupCountNud, 1, 3);

            var flow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.None,
                AutoSize = true,
                Anchor = AnchorStyles.None
            };
            var cancelBtn = new Button { Text = "Cancel", Size = new Size(60, 26) };
            cancelBtn.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
            var okBtn = new Button { Text = "OK", Size = new Size(60, 26) };
            okBtn.Click += (_, _) =>
            {
                CreateBackupOnSave = backupCheckBox.Checked;
                BackupFileCount = (int)(_backupCountNud?.Value ?? 3);
                SelectedFont = _currentFont != null ? (Font)_currentFont.Clone() : null;
                DialogResult = DialogResult.OK;
                Close();
            };
            flow.Controls.Add(cancelBtn);
            flow.Controls.Add(okBtn);
            table.Controls.Add(flow, 0, 4);
            table.SetColumnSpan(flow, 2);

            Controls.Add(table);
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _currentFont != null)
            {
                _currentFont.Dispose();
                _currentFont = null;
            }
            base.Dispose(disposing);
        }
    }

    public class ColorControlsForm : Form
    {
        public Color HeaderColor { get; private set; }
        public Color FrozenColor { get; private set; }
        public Color EvenRowColor { get; private set; }
        public Color OddRowColor { get; private set; }
        public Color SearchHighlightColor { get; private set; }
        public Color ExpansionRowColor { get; private set; }

        public ColorControlsForm(Color headerColor, Color frozenColor, Color evenRowColor, Color oddRowColor, Color searchHighlightColor, Color expansionRowColor)
        {
            HeaderColor = headerColor;
            FrozenColor = frozenColor;
            EvenRowColor = evenRowColor;
            OddRowColor = oddRowColor;
            SearchHighlightColor = searchHighlightColor;
            ExpansionRowColor = expansionRowColor;
            Text = "Color Controls";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(380, 360);
            MinimumSize = new Size(360, 340);
            Padding = new Padding(12);

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(0),
                ColumnCount = 2,
                RowCount = 7,
                Height = 7 * 40
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            AddColorRow(table, 0, "Header Color", headerColor, c => HeaderColor = c);
            AddColorRow(table, 1, "Frozen Color", frozenColor, c => FrozenColor = c);
            AddColorRow(table, 2, "Even Rows", evenRowColor, c => EvenRowColor = c);
            AddColorRow(table, 3, "Odd Rows", oddRowColor, c => OddRowColor = c);
            AddColorRow(table, 4, "Search Highlight", searchHighlightColor, c => SearchHighlightColor = c);
            AddColorRow(table, 5, "Expansion Row", expansionRowColor, c => ExpansionRowColor = c);

            var flow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.None,
                AutoSize = true,
            };
            var cancelBtn = new Button { Text = "Cancel", Size = new Size(85, 30) };
            cancelBtn.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
            var okBtn = new Button { Text = "OK", Size = new Size(85, 30) };
            okBtn.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
            flow.Controls.Add(cancelBtn);
            flow.Controls.Add(okBtn);
            table.Controls.Add(flow, 0, 6);
            table.SetColumnSpan(flow, 2);
            flow.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            Controls.Add(table);
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void AddColorRow(TableLayoutPanel table, int row, string labelText, Color color, Action<Color> setColor)
        {
            var label = new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left };
            var btn = new Button
            {
                BackColor = color,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(120, 28),
                Anchor = AnchorStyles.Left
            };
            btn.FlatAppearance.BorderColor = Color.DarkGray;
            btn.Click += (_, _) =>
            {
                using (var dlg = new ColorDialog { Color = btn.BackColor })
                {
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        btn.BackColor = dlg.Color;
                        setColor(dlg.Color);
                        btn.Invalidate();
                        btn.Update();
                    }
                }
            };
            table.Controls.Add(label, 0, row);
            table.Controls.Add(btn, 1, row);
        }
    }

    public class TextColorControlsForm : Form
    {
        public Color TabTextColor { get; private set; }
        public Color HeaderTextColor { get; private set; }
        public Color CellTextColor { get; private set; }
        public Color CellFrozenTextColor { get; private set; }
        public Color WorkspaceEntryTextColor { get; private set; }
        public Color WorkspaceFilesTextColor { get; private set; }
        public Color EditHistoryTextColor { get; private set; }
        public Color MenuStandbyTextColor { get; private set; }
        public Color MenuActiveTextColor { get; private set; }
        public Color TabTextColorDark { get; private set; }
        public Color HeaderTextColorDark { get; private set; }
        public Color CellTextColorDark { get; private set; }
        public Color CellFrozenTextColorDark { get; private set; }
        public Color WorkspaceEntryTextColorDark { get; private set; }
        public Color WorkspaceFilesTextColorDark { get; private set; }
        public Color EditHistoryTextColorDark { get; private set; }
        public Color MenuStandbyTextColorDark { get; private set; }
        public Color MenuActiveTextColorDark { get; private set; }
        public Color ButtonTextColor { get; private set; }
        public Color ButtonTextColorDark { get; private set; }

        public TextColorControlsForm(Color tabText, Color headerText, Color cellText, Color cellFrozenText, Color workspaceEntryText, Color workspaceFilesText, Color editHistoryText, Color menuStandbyText, Color menuActiveText, Color buttonText, Color tabTextDark, Color headerTextDark, Color cellTextDark, Color cellFrozenTextDark, Color workspaceEntryTextDark, Color workspaceFilesTextDark, Color editHistoryTextDark, Color menuStandbyTextDark, Color menuActiveTextDark, Color buttonTextDark)
        {
            TabTextColor = tabText;
            HeaderTextColor = headerText;
            CellTextColor = cellText;
            CellFrozenTextColor = cellFrozenText;
            WorkspaceEntryTextColor = workspaceEntryText;
            WorkspaceFilesTextColor = workspaceFilesText;
            EditHistoryTextColor = editHistoryText;
            MenuStandbyTextColor = menuStandbyText;
            MenuActiveTextColor = menuActiveText;
            ButtonTextColor = buttonText;
            TabTextColorDark = tabTextDark;
            HeaderTextColorDark = headerTextDark;
            CellTextColorDark = cellTextDark;
            CellFrozenTextColorDark = cellFrozenTextDark;
            WorkspaceEntryTextColorDark = workspaceEntryTextDark;
            WorkspaceFilesTextColorDark = workspaceFilesTextDark;
            EditHistoryTextColorDark = editHistoryTextDark;
            MenuStandbyTextColorDark = menuStandbyTextDark;
            MenuActiveTextColorDark = menuActiveTextDark;
            ButtonTextColorDark = buttonTextDark;
            Text = "Text Color Controls";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            Size = new Size(400, 720);
            MinimumSize = new Size(380, 640);
            Padding = new Padding(12);

            const int rowHeight = 36;
            const int sectionHeaderHeight = 28;

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Padding = new Padding(0),
                ColumnCount = 2,
                RowCount = 23,
                Height = 2 * sectionHeaderHeight + 20 * rowHeight + 50
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, sectionHeaderHeight));
            for (int r = 1; r <= 3; r++) table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight)); // Cell (frozen)
            for (int r = 5; r <= 10; r++) table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, sectionHeaderHeight));
            for (int r = 12; r <= 14; r++) table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight)); // Cell (frozen) dark
            for (int r = 16; r <= 21; r++) table.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            var lightLabel = new Label
            {
                Text = "For Light Mode",
                Font = new Font(Font.FontFamily, 9.5f, FontStyle.Bold),
                Anchor = AnchorStyles.Left,
                ForeColor = Color.DarkSlateGray
            };
            table.Controls.Add(lightLabel, 0, 0);
            table.SetColumnSpan(lightLabel, 2);

            AddColorRow(table, 1, "Tab Text", tabText, c => TabTextColor = c);
            AddColorRow(table, 2, "Header Text", headerText, c => HeaderTextColor = c);
            AddColorRow(table, 3, "Cell Text", cellText, c => CellTextColor = c);
            AddColorRow(table, 4, "Cell (frozen)", cellFrozenText, c => CellFrozenTextColor = c);
            AddColorRow(table, 5, "Workspace Entry Text", workspaceEntryText, c => WorkspaceEntryTextColor = c);
            AddColorRow(table, 6, "Workspace Files Text", workspaceFilesText, c => WorkspaceFilesTextColor = c);
            AddColorRow(table, 7, "Edit History Text", editHistoryText, c => EditHistoryTextColor = c);
            AddColorRow(table, 8, "Menu Standby Text", menuStandbyText, c => MenuStandbyTextColor = c);
            AddColorRow(table, 9, "Menu Active Text", menuActiveText, c => MenuActiveTextColor = c);
            AddColorRow(table, 10, "Button Text", buttonText, c => ButtonTextColor = c);

            var darkLabel = new Label
            {
                Text = "For Dark Mode",
                Font = new Font(Font.FontFamily, 9.5f, FontStyle.Bold),
                Anchor = AnchorStyles.Left,
                ForeColor = Color.DarkSlateGray
            };
            table.Controls.Add(darkLabel, 0, 11);
            table.SetColumnSpan(darkLabel, 2);

            AddColorRow(table, 12, "Tab Text", tabTextDark, c => TabTextColorDark = c);
            AddColorRow(table, 13, "Header Text", headerTextDark, c => HeaderTextColorDark = c);
            AddColorRow(table, 14, "Cell Text", cellTextDark, c => CellTextColorDark = c);
            AddColorRow(table, 15, "Cell (frozen)", cellFrozenTextDark, c => CellFrozenTextColorDark = c);
            AddColorRow(table, 16, "Workspace Entry Text", workspaceEntryTextDark, c => WorkspaceEntryTextColorDark = c);
            AddColorRow(table, 17, "Workspace Files Text", workspaceFilesTextDark, c => WorkspaceFilesTextColorDark = c);
            AddColorRow(table, 18, "Edit History Text", editHistoryTextDark, c => EditHistoryTextColorDark = c);
            AddColorRow(table, 19, "Menu Standby Text", menuStandbyTextDark, c => MenuStandbyTextColorDark = c);
            AddColorRow(table, 20, "Menu Active Text", menuActiveTextDark, c => MenuActiveTextColorDark = c);
            AddColorRow(table, 21, "Button Text", buttonTextDark, c => ButtonTextColorDark = c);

            var cancelBtn = new Button { Text = "Cancel", Size = new Size(85, 30) };
            cancelBtn.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
            var okBtn = new Button { Text = "OK", Size = new Size(85, 30) };
            okBtn.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
            var flow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Right,
                AutoSize = true,
                Padding = new Padding(0, 10, 12, 0),
                WrapContents = false
            };
            flow.Controls.Add(okBtn);
            flow.Controls.Add(cancelBtn);

            var scrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(0)
            };
            scrollPanel.Controls.Add(table);

            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(0) };
            buttonPanel.Controls.Add(flow);

            Controls.Add(scrollPanel);
            Controls.Add(buttonPanel);
            AcceptButton = okBtn;
            CancelButton = cancelBtn;
        }

        private void AddColorRow(TableLayoutPanel table, int row, string labelText, Color color, Action<Color> setColor)
        {
            var label = new Label { Text = labelText, AutoSize = true, Anchor = AnchorStyles.Left };
            var btn = new Button
            {
                BackColor = color,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(120, 28),
                Anchor = AnchorStyles.Left
            };
            btn.FlatAppearance.BorderColor = Color.DarkGray;
            btn.Click += (_, _) =>
            {
                using (var dlg = new ColorDialog { Color = btn.BackColor })
                {
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        btn.BackColor = dlg.Color;
                        setColor(dlg.Color);
                        btn.Invalidate();
                        btn.Update();
                    }
                }
            };
            table.Controls.Add(label, 0, row);
            table.Controls.Add(btn, 1, row);
        }
    }

    internal sealed class MenuTextColorRenderer : ToolStripProfessionalRenderer
    {
        private readonly Color _activeTextColor;

        public MenuTextColorRenderer(Color activeTextColor)
        {
            _activeTextColor = activeTextColor;
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            if (e.Item.Selected || e.Item.Pressed)
                e.TextColor = _activeTextColor;
            base.OnRenderItemText(e);
        }
    }
}
