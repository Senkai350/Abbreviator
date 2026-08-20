using System;
using System.Runtime.InteropServices;

namespace Abbreviator.Interop
{
    // ВАЖНО про атрибуты в этом файле.
    //
    // Интерфейсы объявлены БЕЗ [ComImport]. ComImport означает «интерфейс
    // реализован COM-объектом, а мы его вызываем»; для такого интерфейса CLR не
    // умеет построить ITypeInfo, и IDispatch::GetIDsOfNames по нему падает.
    // Здесь всё наоборот: интерфейсы реализует управляемый класс, а вызывает их
    // Word. Поэтому нужны обычные управляемые интерфейсы с [Guid] и
    // InterfaceIsDual — ровно так они объявлены в Extensibility.dll и office.dll.
    //
    // [ComVisible(true)] обязателен на каждом типе: сборка помечена
    // ComVisible(false), а типы из сигнатур публичных методов класса Connect
    // (он AutoDual) должны быть видимы COM, иначе CCW для него не создаётся
    // и Word сообщает «произошла ошибка с надстройкой».

    /// <summary>Режим подключения надстройки (Extensibility.ext_ConnectMode).</summary>
    [ComVisible(true)]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4
    }

    /// <summary>Режим отключения надстройки (Extensibility.ext_DisconnectMode).</summary>
    [ComVisible(true)]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    /// <summary>
    /// Интерфейс COM-надстройки Office. Объявлен вручную, чтобы не тянуть
    /// Extensibility.dll. Порядок методов менять нельзя — он задаёт vtable.
    /// </summary>
    [ComVisible(true)]
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
    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }
}
