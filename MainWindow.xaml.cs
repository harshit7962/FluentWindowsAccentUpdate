using Microsoft.Win32;
using System.Drawing;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System;
using System.Dynamic;
using System.Windows.Media.Media3D;
using System.Runtime.CompilerServices;

namespace FluentWindowsAccentUpdate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Win32WindowSettings obj = new Win32WindowSettings();

            var color1 = obj.AccentColor;
            var color2 = obj.AccentColor1;
            var color3 = obj.AccentColor2;
            var color4 = obj.AccentColor3;
            var color5 = obj.AccentColor4;
            var color6 = obj.AccentColor5;
            var color7 = obj.AccentColor6;
        }
    }

    internal static unsafe partial class WinRT
    {
        /// <include file='WinRT.xml' path='doc/member[@name="WinRT.WindowsCreateString"]/*' />
        [DllImport("combase", ExactSpelling = true)]
        public static extern HRESULT WindowsCreateString(ushort* sourceString, uint length, HSTRING* @string);

        /// <include file='WinRT.xml' path='doc/member[@name="WinRT.WindowsDeleteString"]/*' />
        [DllImport("combase", ExactSpelling = true)]
        public static extern HRESULT WindowsDeleteString(HSTRING @string);

        /// <include file='WinRT.xml' path='doc/member[@name="WinRT.RoActivateInstance"]/*' />
        [DllImport("combase", ExactSpelling = true)]
        public static extern HRESULT RoActivateInstance(HSTRING activatableClassId, IInspectable** instance);
    }

    internal readonly unsafe partial struct HSTRING
    {
        public readonly void* Value;

        public int CompareTo(object? obj)
        {
            if (obj is HSTRING other)
            {
                return CompareTo(other);
            }

            return (obj is null) ? 1 : throw new ArgumentException("obj is not an instance of HSTRING.");
        }

        public int CompareTo(HSTRING other) => ((nuint)(Value)).CompareTo((nuint)(other.Value));

        public override bool Equals(object? obj) => (obj is HSTRING other) && Equals(other);

        public bool Equals(HSTRING other) => ((nuint)(Value)).Equals((nuint)(other.Value));

        public override int GetHashCode() => ((nuint)(Value)).GetHashCode();

        public override string ToString() => ((nuint)(Value)).ToString((sizeof(nint) == 4) ? "X8" : "X16");

        public string ToString(string? format, IFormatProvider? formatProvider) => ((nuint)(Value)).ToString(format, formatProvider);
    }

    internal readonly unsafe partial struct HRESULT
    {
        public readonly int Value;

        public bool FAILED => false;

        public HRESULT(int value)
        {
            Value = value;
        }

        public static bool operator ==(HRESULT left, HRESULT right) => left.Value == right.Value;

        public static bool operator !=(HRESULT left, HRESULT right) => left.Value != right.Value;

        public static bool operator <(HRESULT left, HRESULT right) => left.Value < right.Value;

        public static bool operator <=(HRESULT left, HRESULT right) => left.Value <= right.Value;

        public static bool operator >(HRESULT left, HRESULT right) => left.Value > right.Value;

        public static bool operator >=(HRESULT left, HRESULT right) => left.Value >= right.Value;

        public static implicit operator HRESULT(int value) => new HRESULT(value);

        public static implicit operator int(HRESULT value) => value.Value;

        public int CompareTo(object? obj)
        {
            if (obj is HRESULT other)
            {
                return CompareTo(other);
            }

            return (obj is null) ? 1 : throw new ArgumentException("obj is not an instance of HRESULT.");
        }

        public int CompareTo(HRESULT other) => Value.CompareTo(other.Value);

        public override bool Equals(object? obj) => (obj is HRESULT other) && Equals(other);

        public bool Equals(HRESULT other) => Value.Equals(other.Value);

        public override int GetHashCode() => Value.GetHashCode();

        public override string ToString() => Value.ToString("X8");

        public string ToString(string? format, IFormatProvider? formatProvider) => Value.ToString(format, formatProvider);
    }

    internal unsafe class Win32WindowSettings : WindowSettingsImpl
    {
        private readonly IUISettings3* _uiSettings3;

        public Win32WindowSettings()
        {
            const string uiSettingsClassName = "Windows.UI.ViewManagement.UISettings";
            HSTRING hstring;
            fixed (char* name = uiSettingsClassName)
            {
                ThrowIfFailed(WinRT.WindowsCreateString((ushort*)name, (uint)uiSettingsClassName.Length, &hstring));
            }

            IInspectable* inspectable;
            ThrowIfFailed(WinRT.RoActivateInstance(hstring, &inspectable));
            WinRT.WindowsDeleteString(hstring);

            var guid = new Guid("03021BE4-5254-4781-8194-5168F7D06D7B");
            IUISettings3* uiSettings3;
            ThrowIfFailed(inspectable->QueryInterface(&guid, (void**)&uiSettings3));
            _uiSettings3 = uiSettings3;
        }

        public override Color AccentColor => _uiSettings3->GetColorValue(UIColorType.Accent);
        public override Color AccentColor1 => _uiSettings3->GetColorValue(UIColorType.AccentDark1);
        public override Color AccentColor2 => _uiSettings3->GetColorValue(UIColorType.AccentDark2);
        public override Color AccentColor3 => _uiSettings3->GetColorValue(UIColorType.AccentDark3);
        public override Color AccentColor4 => _uiSettings3->GetColorValue(UIColorType.AccentLight1);
        public override Color AccentColor5 => _uiSettings3->GetColorValue(UIColorType.AccentLight2);
        public override Color AccentColor6 => _uiSettings3->GetColorValue(UIColorType.AccentLight3);

        // Extract of the IUISettings3 from windows.ui.viewmanagement.idl
        private struct IUISettings3
        {
            public void** lpVtbl;

            public Color GetColorValue(UIColorType desiredColor)
            {
                UIColor value;
                // The GetColorValue method comes right after IInspectable and is at VTBL slot 6
                ((delegate* unmanaged<IUISettings3*, UIColorType, UIColor*, int>)(lpVtbl[6]))((IUISettings3*)Unsafe.AsPointer(ref this), desiredColor, &value);
                return Color.FromArgb(*(int*)&value);
            }
        }

        private enum UIColorType
        {
            Background = 0,
            Foreground = 1,
            AccentDark3 = 2,
            AccentDark2 = 3,
            AccentDark1 = 4,
            Accent = 5,
            AccentLight1 = 6,
            AccentLight2 = 7,
            AccentLight3 = 8,
            Complement = 9
        };

        private readonly record struct UIColor(byte A, byte R, byte G, byte B);

        private static void ThrowIfFailed(HRESULT result)
        {
            if (result.FAILED)
            {
                throw new PlatformNotSupportedException("Unable to access `Windows.UI.ViewManagement.UISettings`. Platform not supported");
            }
        }
    }

    internal abstract class WindowSettingsImpl
    {
        public abstract Color AccentColor { get; }
        public abstract Color AccentColor1 { get; }
        public abstract Color AccentColor2 { get; }
        public abstract Color AccentColor3 { get; }
        public abstract Color AccentColor4 { get; }
        public abstract Color AccentColor5 { get; }
        public abstract Color AccentColor6 { get; }
    }

    [Guid("AF86E2E0-B12D-4C6A-9C5A-D7AA65101E90")]
    internal unsafe partial struct IInspectable
    {

        public void** lpVtbl;

        /// <inheritdoc cref="IUnknown.QueryInterface" />
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public HRESULT QueryInterface(Guid* riid, void** ppvObject)
        {
            return ((delegate* unmanaged<IInspectable*, Guid*, void**, int>)(lpVtbl[0]))((IInspectable*)Unsafe.AsPointer(ref this), riid, ppvObject);
        }
    }
}