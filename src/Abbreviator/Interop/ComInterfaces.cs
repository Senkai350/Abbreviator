using System;
using System.Runtime.InteropServices;

namespace Abbreviator.Interop
{
    /// <summary>Режим подключения надстройки (Extensibility.ext_ConnectMode).</summary>
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4
    }

    /// <summary>Режим отключения надстройки (Extensibility.ext_DisconnectMode).</summary>
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    /// <summary>
    /// Интерфейс COM-надстройки Office. Объявлен вручную, чтобы не тянуть Extensibility.dll.
    /// Порядок методов менять нельзя — он совпадает с vtable оригинального интерфейса.
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            ext_ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }

    /// <summary>
    /// Office.IRibbonExtensibility. Объявлен вручную, чтобы не зависеть от Office PIA.
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }
}
