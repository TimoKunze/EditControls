// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include <initguid.h>

#import <libid:A8F9B8E7-E699-4fce-A647-72C877F8E632> version("1.10") raw_dispinterfaces
#import <libid:EA57D88C-8144-415a-9666-B7067B74C295> version("1.10") raw_dispinterfaces

DEFINE_GUID(LIBID_EditCtlsLibU, 0xA8F9B8E7, 0xE699, 0x4FCE, 0xA6, 0x47, 0x72, 0xC8, 0x77, 0xF8, 0xE6, 0x32);
DEFINE_GUID(LIBID_EditCtlsLibA, 0xEA57D88C, 0x8144, 0x415A, 0x96, 0x66, 0xB7, 0x06, 0x7B, 0x74, 0xC2, 0x95);

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_TXTBOX0U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_TXTBOX1U, CMainDlg, &__uuidof(EditCtlsLibU::_ITextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_IPADDRU, CMainDlg, &__uuidof(EditCtlsLibU::_IIPAddressBoxEvents), &LIBID_EditCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_HKBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IHotKeyBoxEvents), &LIBID_EditCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_UDTXTBOXU, CMainDlg, &__uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibU, 1, 10>,
    public IDispEventImpl<IDC_TXTBOX0A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_TXTBOX1A, CMainDlg, &__uuidof(EditCtlsLibA::_ITextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_IPADDRA, CMainDlg, &__uuidof(EditCtlsLibA::_IIPAddressBoxEvents), &LIBID_EditCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_HKBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IHotKeyBoxEvents), &LIBID_EditCtlsLibA, 1, 10>,
    public IDispEventImpl<IDC_UDTXTBOXA, CMainDlg, &__uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), &LIBID_EditCtlsLibA, 1, 10>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CButton> aboutButton;
	CContainedWindowT<CAxWindow> txtbox0UWnd;
	CContainedWindowT<CAxWindow> txtbox1UWnd;
	CContainedWindowT<CAxWindow> ipaddrUWnd;
	CContainedWindowT<CAxWindow> hkboxUWnd;
	CContainedWindowT<CAxWindow> udtxtboxUWnd;
	CContainedWindowT<CAxWindow> txtbox0AWnd;
	CContainedWindowT<CAxWindow> txtbox1AWnd;
	CContainedWindowT<CAxWindow> ipaddrAWnd;
	CContainedWindowT<CAxWindow> hkboxAWnd;
	CContainedWindowT<CAxWindow> udtxtboxAWnd;

	CMainDlg() :
	    aboutButton(this, 1),
	    txtbox0UWnd(this, 2),
	    txtbox1UWnd(this, 3),
	    ipaddrUWnd(this, 4),
	    hkboxUWnd(this, 5),
	    udtxtboxUWnd(this, 6),
	    txtbox0AWnd(this, 7),
	    txtbox1AWnd(this, 8),
	    ipaddrAWnd(this, 9),
	    hkboxAWnd(this, 10),
	    udtxtboxAWnd(this, 11)
	{
		hFocusedControl = NULL;
	}

	HWND hFocusedControl;

	struct Controls
	{
		CEdit logEdit;
		CComPtr<EditCtlsLibU::ITextBox> txtbox0U;
		CComPtr<EditCtlsLibU::ITextBox> txtbox1U;
		CComPtr<EditCtlsLibU::IIPAddressBox> ipaddrU;
		CComPtr<EditCtlsLibU::IHotKeyBox> hkboxU;
		CComPtr<EditCtlsLibU::IUpDownTextBox> udtxtboxU;
		CComPtr<EditCtlsLibA::ITextBox> txtbox0A;
		CComPtr<EditCtlsLibA::ITextBox> txtbox1A;
		CComPtr<EditCtlsLibA::IIPAddressBox> ipaddrA;
		CComPtr<EditCtlsLibA::IHotKeyBox> hkboxA;
		CComPtr<EditCtlsLibA::IUpDownTextBox> udtxtboxA;
	} controls;

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)
		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocusAbout)

		ALT_MSG_MAP(2)
		ALT_MSG_MAP(3)
		ALT_MSG_MAP(4)
		ALT_MSG_MAP(5)
		ALT_MSG_MAP(6)
		ALT_MSG_MAP(7)
		ALT_MSG_MAP(8)
		ALT_MSG_MAP(9)
		ALT_MSG_MAP(10)
		ALT_MSG_MAP(11)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 0, TextChangedTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 1, AbortedDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 2, BeforeDrawTextTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 3, BeginDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 4, BeginRDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 5, ClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 6, ContextMenuTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 7, DblClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 8, DestroyedControlWindowTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 9, DragMouseMoveTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 10, DropTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 11, KeyDownTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 12, KeyPressTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 13, KeyUpTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 14, MClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 15, MDblClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 16, MouseDownTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 17, MouseEnterTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 18, MouseHoverTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 19, MouseLeaveTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 20, MouseMoveTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 21, MouseUpTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 22, OLECompleteDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 23, OLEDragDropTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 24, OLEDragEnterTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 25, OLEDragLeaveTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 26, OLEDragMouseMoveTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 27, OLEGiveFeedbackTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 28, OLEQueryContinueDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 29, OLESetDataTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 30, OLEStartDragTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 31, OutOfMemoryTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 32, RClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 33, RDblClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 34, RecreatedControlWindowTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 35, ResizedControlWindowTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 36, ScrollingTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 37, TruncatedTextTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 38, WritingDirectionChangedTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 39, OLEDragEnterPotentialTargetTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 40, OLEDragLeavePotentialTargetTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 41, OLEReceivedNewDataTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 42, MouseWheelTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 43, XClickTxtbox0U)
		SINK_ENTRY_EX(IDC_TXTBOX0U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 44, XDblClickTxtbox0U)

		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 0, TextChangedTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 1, AbortedDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 2, BeforeDrawTextTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 3, BeginDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 4, BeginRDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 5, ClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 6, ContextMenuTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 7, DblClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 8, DestroyedControlWindowTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 9, DragMouseMoveTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 10, DropTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 11, KeyDownTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 12, KeyPressTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 13, KeyUpTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 14, MClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 15, MDblClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 16, MouseDownTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 17, MouseEnterTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 18, MouseHoverTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 19, MouseLeaveTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 20, MouseMoveTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 21, MouseUpTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 22, OLECompleteDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 23, OLEDragDropTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 24, OLEDragEnterTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 25, OLEDragLeaveTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 26, OLEDragMouseMoveTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 27, OLEGiveFeedbackTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 28, OLEQueryContinueDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 29, OLESetDataTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 30, OLEStartDragTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 31, OutOfMemoryTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 32, RClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 33, RDblClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 34, RecreatedControlWindowTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 35, ResizedControlWindowTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 36, ScrollingTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 37, TruncatedTextTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 38, WritingDirectionChangedTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 39, OLEDragEnterPotentialTargetTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 40, OLEDragLeavePotentialTargetTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 41, OLEReceivedNewDataTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 42, MouseWheelTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 43, XClickTxtbox1U)
		SINK_ENTRY_EX(IDC_TXTBOX1U, __uuidof(EditCtlsLibU::_ITextBoxEvents), 44, XDblClickTxtbox1U)

		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 0, AddressChangedIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 1, ClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 2, ContextMenuIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 3, DblClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 4, DestroyedControlWindowIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 5, FieldTextChangedIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 6, KeyDownIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 7, KeyPressIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 8, KeyUpIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 9, MClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 10, MDblClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 11, MouseDownIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 12, MouseEnterIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 13, MouseHoverIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 14, MouseLeaveIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 15, MouseMoveIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 16, MouseUpIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 17, OLEDragDropIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 18, OLEDragEnterIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 19, OLEDragLeaveIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 20, OLEDragMouseMoveIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 21, RClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 22, RDblClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 23, RecreatedControlWindowIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 24, ResizedControlWindowIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 25, MouseWheelIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 26, XClickIpaddrU)
		SINK_ENTRY_EX(IDC_IPADDRU, __uuidof(EditCtlsLibU::_IIPAddressBoxEvents), 27, XDblClickIpaddrU)

		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 0, KeyDownHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 1, ClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 2, ContextMenuHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 3, DblClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 4, DestroyedControlWindowHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 5, KeyPressHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 6, KeyUpHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 7, MClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 8, MDblClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 9, MouseDownHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 10, MouseEnterHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 11, MouseHoverHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 12, MouseLeaveHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 13, MouseMoveHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 14, MouseUpHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 15, OLEDragDropHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 16, OLEDragEnterHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 17, OLEDragLeaveHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 18, OLEDragMouseMoveHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 19, RClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 20, RDblClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 21, RecreatedControlWindowHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 22, ResizedControlWindowHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 23, MouseWheelHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 24, XClickHkboxU)
		SINK_ENTRY_EX(IDC_HKBOXU, __uuidof(EditCtlsLibU::_IHotKeyBoxEvents), 25, XDblClickHkboxU)

		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 0, ValueChangedUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 1, BeforeDrawTextUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 2, ChangedAcceleratorsUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 3, ChangingAcceleratorsUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 4, ClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 5, ContextMenuUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 6, DblClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 7, DestroyedControlWindowUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 8, KeyDownUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 9, KeyPressUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 10, KeyUpUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 11, MClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 12, MDblClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 13, MouseDownUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 14, MouseEnterUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 15, MouseHoverUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 16, MouseLeaveUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 17, MouseMoveUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 18, MouseUpUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 19, OLEDragDropUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 20, OLEDragEnterUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 21, OLEDragLeaveUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 22, OLEDragMouseMoveUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 23, OutOfMemoryUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 24, RClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 25, RDblClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 26, RecreatedControlWindowUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 27, ResizedControlWindowUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 28, TextChangedUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 29, TruncatedTextUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 30, ValueChangingUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 31, WritingDirectionChangedUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 32, MouseWheelUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 33, XClickUdtxtboxU)
		SINK_ENTRY_EX(IDC_UDTXTBOXU, __uuidof(EditCtlsLibU::_IUpDownTextBoxEvents), 34, XDblClickUdtxtboxU)

		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 0, TextChangedTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 1, AbortedDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 2, BeforeDrawTextTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 3, BeginDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 4, BeginRDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 5, ClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 6, ContextMenuTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 7, DblClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 8, DestroyedControlWindowTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 9, DragMouseMoveTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 10, DropTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 11, KeyDownTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 12, KeyPressTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 13, KeyUpTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 14, MClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 15, MDblClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 16, MouseDownTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 17, MouseEnterTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 18, MouseHoverTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 19, MouseLeaveTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 20, MouseMoveTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 21, MouseUpTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 22, OLECompleteDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 23, OLEDragDropTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 24, OLEDragEnterTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 25, OLEDragLeaveTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 26, OLEDragMouseMoveTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 27, OLEGiveFeedbackTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 28, OLEQueryContinueDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 29, OLESetDataTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 30, OLEStartDragTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 31, OutOfMemoryTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 32, RClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 33, RDblClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 34, RecreatedControlWindowTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 35, ResizedControlWindowTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 36, ScrollingTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 37, TruncatedTextTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 38, WritingDirectionChangedTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 39, OLEDragEnterPotentialTargetTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 40, OLEDragLeavePotentialTargetTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 41, OLEReceivedNewDataTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 42, MouseWheelTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 43, XClickTxtbox0A)
		SINK_ENTRY_EX(IDC_TXTBOX0A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 44, XDblClickTxtbox0A)

		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 0, TextChangedTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 1, AbortedDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 2, BeforeDrawTextTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 3, BeginDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 4, BeginRDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 5, ClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 6, ContextMenuTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 7, DblClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 8, DestroyedControlWindowTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 9, DragMouseMoveTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 10, DropTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 11, KeyDownTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 12, KeyPressTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 13, KeyUpTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 14, MClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 15, MDblClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 16, MouseDownTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 17, MouseEnterTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 18, MouseHoverTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 19, MouseLeaveTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 20, MouseMoveTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 21, MouseUpTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 22, OLECompleteDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 23, OLEDragDropTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 24, OLEDragEnterTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 25, OLEDragLeaveTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 26, OLEDragMouseMoveTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 27, OLEGiveFeedbackTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 28, OLEQueryContinueDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 29, OLESetDataTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 30, OLEStartDragTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 31, OutOfMemoryTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 32, RClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 33, RDblClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 34, RecreatedControlWindowTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 35, ResizedControlWindowTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 36, ScrollingTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 37, TruncatedTextTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 38, WritingDirectionChangedTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 39, OLEDragEnterPotentialTargetTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 40, OLEDragLeavePotentialTargetTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 41, OLEReceivedNewDataTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 42, MouseWheelTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 43, XClickTxtbox1A)
		SINK_ENTRY_EX(IDC_TXTBOX1A, __uuidof(EditCtlsLibA::_ITextBoxEvents), 44, XDblClickTxtbox1A)

		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 0, AddressChangedIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 1, ClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 2, ContextMenuIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 3, DblClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 4, DestroyedControlWindowIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 5, FieldTextChangedIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 6, KeyDownIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 7, KeyPressIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 8, KeyUpIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 9, MClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 10, MDblClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 11, MouseDownIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 12, MouseEnterIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 13, MouseHoverIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 14, MouseLeaveIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 15, MouseMoveIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 16, MouseUpIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 17, OLEDragDropIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 18, OLEDragEnterIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 19, OLEDragLeaveIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 20, OLEDragMouseMoveIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 21, RClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 22, RDblClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 23, RecreatedControlWindowIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 24, ResizedControlWindowIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 25, MouseWheelIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 26, XClickIpaddrA)
		SINK_ENTRY_EX(IDC_IPADDRA, __uuidof(EditCtlsLibA::_IIPAddressBoxEvents), 27, XDblClickIpaddrA)

		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 0, KeyDownHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 1, ClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 2, ContextMenuHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 3, DblClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 4, DestroyedControlWindowHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 5, KeyPressHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 6, KeyUpHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 7, MClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 8, MDblClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 9, MouseDownHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 10, MouseEnterHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 11, MouseHoverHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 12, MouseLeaveHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 13, MouseMoveHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 14, MouseUpHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 15, OLEDragDropHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 16, OLEDragEnterHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 17, OLEDragLeaveHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 18, OLEDragMouseMoveHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 19, RClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 20, RDblClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 21, RecreatedControlWindowHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 22, ResizedControlWindowHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 23, MouseWheelHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 24, XClickHkboxA)
		SINK_ENTRY_EX(IDC_HKBOXA, __uuidof(EditCtlsLibA::_IHotKeyBoxEvents), 25, XDblClickHkboxA)

		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 0, ValueChangedUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 1, BeforeDrawTextUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 2, ChangedAcceleratorsUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 3, ChangingAcceleratorsUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 4, ClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 5, ContextMenuUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 6, DblClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 7, DestroyedControlWindowUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 8, KeyDownUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 9, KeyPressUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 10, KeyUpUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 11, MClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 12, MDblClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 13, MouseDownUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 14, MouseEnterUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 15, MouseHoverUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 16, MouseLeaveUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 17, MouseMoveUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 18, MouseUpUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 19, OLEDragDropUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 20, OLEDragEnterUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 21, OLEDragLeaveUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 22, OLEDragMouseMoveUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 23, OutOfMemoryUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 24, RClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 25, RDblClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 26, RecreatedControlWindowUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 27, ResizedControlWindowUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 28, TextChangedUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 29, TruncatedTextUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 30, ValueChangingUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 31, WritingDirectionChangedUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 32, MouseWheelUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 33, XClickUdtxtboxA)
		SINK_ENTRY_EX(IDC_UDTXTBOXA, __uuidof(EditCtlsLibA::_IUpDownTextBoxEvents), 34, XDblClickUdtxtboxA)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnSetFocusAbout(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& bHandled);

	void AddLogEntry(CAtlString text);
	void CloseDialog(int nVal);

	void __stdcall AbortedDragTxtbox0U();
	void __stdcall BeforeDrawTextTxtbox0U();
	void __stdcall BeginDragTxtbox0U(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall BeginRDragTxtbox0U(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall ClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall ContextMenuTxtbox0U(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTxtbox0U(long hWnd);
	void __stdcall DragMouseMoveTxtbox0U(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropTxtbox0U(short button, short shift, long x, long y);
	void __stdcall KeyDownTxtbox0U(short* keyCode, short shift);
	void __stdcall KeyPressTxtbox0U(short* keyAscii);
	void __stdcall KeyUpTxtbox0U(short* keyCode, short shift);
	void __stdcall MClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MDblClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseDownTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseEnterTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseHoverTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseMoveTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseUpTxtbox0U(short button, short shift, long x, long y);
	void __stdcall MouseWheelTxtbox0U(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTxtbox0U(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTxtbox0U(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragLeavePotentialTargetTxtbox0U();
	void __stdcall OLEDragMouseMoveTxtbox0U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackTxtbox0U(EditCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTxtbox0U(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTxtbox0U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTxtbox0U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTxtbox0U(LPDISPATCH data);
	void __stdcall OutOfMemoryTxtbox0U();
	void __stdcall RClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall RDblClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTxtbox0U(long hWnd);
	void __stdcall ResizedControlWindowTxtbox0U();
	void __stdcall ScrollingTxtbox0U(EditCtlsLibU::ScrollAxisConstants axis);
	void __stdcall TextChangedTxtbox0U();
	void __stdcall TruncatedTextTxtbox0U();
	void __stdcall WritingDirectionChangedTxtbox0U(EditCtlsLibU::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickTxtbox0U(short button, short shift, long x, long y);
	void __stdcall XDblClickTxtbox0U(short button, short shift, long x, long y);

	void __stdcall AbortedDragTxtbox1U();
	void __stdcall BeforeDrawTextTxtbox1U();
	void __stdcall BeginDragTxtbox1U(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall BeginRDragTxtbox1U(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall ClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall ContextMenuTxtbox1U(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTxtbox1U(long hWnd);
	void __stdcall DragMouseMoveTxtbox1U(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropTxtbox1U(short button, short shift, long x, long y);
	void __stdcall KeyDownTxtbox1U(short* keyCode, short shift);
	void __stdcall KeyPressTxtbox1U(short* keyAscii);
	void __stdcall KeyUpTxtbox1U(short* keyCode, short shift);
	void __stdcall MClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MDblClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseDownTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseEnterTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseHoverTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseMoveTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseUpTxtbox1U(short button, short shift, long x, long y);
	void __stdcall MouseWheelTxtbox1U(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTxtbox1U(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTxtbox1U(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragLeavePotentialTargetTxtbox1U();
	void __stdcall OLEDragMouseMoveTxtbox1U(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackTxtbox1U(EditCtlsLibU::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTxtbox1U(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibU::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTxtbox1U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTxtbox1U(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTxtbox1U(LPDISPATCH data);
	void __stdcall OutOfMemoryTxtbox1U();
	void __stdcall RClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall RDblClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTxtbox1U(long hWnd);
	void __stdcall ResizedControlWindowTxtbox1U();
	void __stdcall ScrollingTxtbox1U(EditCtlsLibU::ScrollAxisConstants axis);
	void __stdcall TextChangedTxtbox1U();
	void __stdcall TruncatedTextTxtbox1U();
	void __stdcall WritingDirectionChangedTxtbox1U(EditCtlsLibU::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickTxtbox1U(short button, short shift, long x, long y);
	void __stdcall XDblClickTxtbox1U(short button, short shift, long x, long y);

	void __stdcall AddressChangedIpaddrU(long editBoxIndex, long* newFieldValue);
	void __stdcall ClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall ContextMenuIpaddrU(long editBoxIndex, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowIpaddrU(long hWnd);
	void __stdcall FieldTextChangedIpaddrU(long editBoxIndex);
	void __stdcall KeyDownIpaddrU(long editBoxIndex, short* keyCode, short shift);
	void __stdcall KeyPressIpaddrU(long editBoxIndex, short* keyAscii);
	void __stdcall KeyUpIpaddrU(long editBoxIndex, short* keyCode, short shift);
	void __stdcall MClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseDownIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseEnterIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseHoverIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseLeaveIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseMoveIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseUpIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseWheelIpaddrU(long editBoxIndex, short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveIpaddrU(long editBoxIndex, LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveIpaddrU(long editBoxIndex, LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall RClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall RDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowIpaddrU(long hWnd);
	void __stdcall ResizedControlWindowIpaddrU();
	void __stdcall XClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall XDblClickIpaddrU(long editBoxIndex, short button, short shift, long x, long y);

	void __stdcall ClickHkboxU(short button, short shift, long x, long y);
	void __stdcall ContextMenuHkboxU(short button, short shift, long x, long y);
	void __stdcall DblClickHkboxU(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowHkboxU(long hWnd);
	void __stdcall KeyDownHkboxU(short* keyCode, short shift);
	void __stdcall KeyPressHkboxU(short* keyAscii);
	void __stdcall KeyUpHkboxU(short* keyCode, short shift);
	void __stdcall MClickHkboxU(short button, short shift, long x, long y);
	void __stdcall MDblClickHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseDownHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseEnterHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseHoverHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseLeaveHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseMoveHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseUpHkboxU(short button, short shift, long x, long y);
	void __stdcall MouseWheelHkboxU(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveHkboxU(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveHkboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall RClickHkboxU(short button, short shift, long x, long y);
	void __stdcall RDblClickHkboxU(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowHkboxU(long hWnd);
	void __stdcall ResizedControlWindowHkboxU();
	void __stdcall XClickHkboxU(short button, short shift, long x, long y);
	void __stdcall XDblClickHkboxU(short button, short shift, long x, long y);

	void __stdcall BeforeDrawTextUdtxtboxU();
	void __stdcall ChangedAcceleratorsUdtxtboxU(LPDISPATCH accelerators);
	void __stdcall ChangingAcceleratorsUdtxtboxU(LPDISPATCH accelerators, VARIANT_BOOL* cancelChanges);
	void __stdcall ClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowUdtxtboxU(long hWnd);
	void __stdcall KeyDownUdtxtboxU(short* keyCode, short shift);
	void __stdcall KeyPressUdtxtboxU(short* keyAscii);
	void __stdcall KeyUpUdtxtboxU(short* keyCode, short shift);
	void __stdcall MClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseDownUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseUpUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeaveUdtxtboxU(LPDISPATCH data, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveUdtxtboxU(LPDISPATCH data, EditCtlsLibU::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall OutOfMemoryUdtxtboxU();
	void __stdcall RClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowUdtxtboxU(long hWnd);
	void __stdcall ResizedControlWindowUdtxtboxU();
	void __stdcall TextChangedUdtxtboxU();
	void __stdcall TruncatedTextUdtxtboxU();
	void __stdcall ValueChangedUdtxtboxU();
	void __stdcall ValueChangingUdtxtboxU(long currentValue, long delta, VARIANT_BOOL* cancelChange);
	void __stdcall WritingDirectionChangedUdtxtboxU(EditCtlsLibU::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);
	void __stdcall XDblClickUdtxtboxU(short button, short shift, long x, long y, EditCtlsLibU::HitTestConstants hitTestDetails);

	void __stdcall AbortedDragTxtbox0A();
	void __stdcall BeforeDrawTextTxtbox0A();
	void __stdcall BeginDragTxtbox0A(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall BeginRDragTxtbox0A(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall ClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall ContextMenuTxtbox0A(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTxtbox0A(long hWnd);
	void __stdcall DragMouseMoveTxtbox0A(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropTxtbox0A(short button, short shift, long x, long y);
	void __stdcall KeyDownTxtbox0A(short* keyCode, short shift);
	void __stdcall KeyPressTxtbox0A(short* keyAscii);
	void __stdcall KeyUpTxtbox0A(short* keyCode, short shift);
	void __stdcall MClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MDblClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseDownTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseEnterTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseHoverTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseMoveTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseUpTxtbox0A(short button, short shift, long x, long y);
	void __stdcall MouseWheelTxtbox0A(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTxtbox0A(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTxtbox0A(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragLeavePotentialTargetTxtbox0A();
	void __stdcall OLEDragMouseMoveTxtbox0A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackTxtbox0A(EditCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTxtbox0A(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTxtbox0A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTxtbox0A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTxtbox0A(LPDISPATCH data);
	void __stdcall OutOfMemoryTxtbox0A();
	void __stdcall RClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall RDblClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTxtbox0A(long hWnd);
	void __stdcall ResizedControlWindowTxtbox0A();
	void __stdcall ScrollingTxtbox0A(EditCtlsLibA::ScrollAxisConstants axis);
	void __stdcall TextChangedTxtbox0A();
	void __stdcall TruncatedTextTxtbox0A();
	void __stdcall WritingDirectionChangedTxtbox0A(EditCtlsLibA::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickTxtbox0A(short button, short shift, long x, long y);
	void __stdcall XDblClickTxtbox0A(short button, short shift, long x, long y);

	void __stdcall AbortedDragTxtbox1A();
	void __stdcall BeforeDrawTextTxtbox1A();
	void __stdcall BeginDragTxtbox1A(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall BeginRDragTxtbox1A(long firstChar, long lastChar, short button, short shift, long x, long y);
	void __stdcall ClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall ContextMenuTxtbox1A(short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowTxtbox1A(long hWnd);
	void __stdcall DragMouseMoveTxtbox1A(short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall DropTxtbox1A(short button, short shift, long x, long y);
	void __stdcall KeyDownTxtbox1A(short* keyCode, short shift);
	void __stdcall KeyPressTxtbox1A(short* keyAscii);
	void __stdcall KeyUpTxtbox1A(short* keyCode, short shift);
	void __stdcall MClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MDblClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseDownTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseEnterTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseHoverTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseLeaveTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseMoveTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseUpTxtbox1A(short button, short shift, long x, long y);
	void __stdcall MouseWheelTxtbox1A(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLECompleteDragTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants performedEffect);
	void __stdcall OLEDragDropTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEDragEnterPotentialTargetTxtbox1A(long hWndPotentialTarget);
	void __stdcall OLEDragLeaveTxtbox1A(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragLeavePotentialTargetTxtbox1A();
	void __stdcall OLEDragMouseMoveTxtbox1A(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, long* autoHScrollVelocity, long* autoVScrollVelocity);
	void __stdcall OLEGiveFeedbackTxtbox1A(EditCtlsLibA::OLEDropEffectConstants effect, VARIANT_BOOL* useDefaultCursors);
	void __stdcall OLEQueryContinueDragTxtbox1A(VARIANT_BOOL pressedEscape, short button, short shift, EditCtlsLibA::OLEActionToContinueWithConstants* actionToContinueWith);
	void __stdcall OLEReceivedNewDataTxtbox1A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLESetDataTxtbox1A(LPDISPATCH data, long formatID, long index, long dataOrViewAspect);
	void __stdcall OLEStartDragTxtbox1A(LPDISPATCH data);
	void __stdcall OutOfMemoryTxtbox1A();
	void __stdcall RClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall RDblClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowTxtbox1A(long hWnd);
	void __stdcall ResizedControlWindowTxtbox1A();
	void __stdcall ScrollingTxtbox1A(EditCtlsLibA::ScrollAxisConstants axis);
	void __stdcall TextChangedTxtbox1A();
	void __stdcall TruncatedTextTxtbox1A();
	void __stdcall WritingDirectionChangedTxtbox1A(EditCtlsLibA::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickTxtbox1A(short button, short shift, long x, long y);
	void __stdcall XDblClickTxtbox1A(short button, short shift, long x, long y);

	void __stdcall AddressChangedIpaddrA(long editBoxIndex, long* newFieldValue);
	void __stdcall ClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall ContextMenuIpaddrA(long editBoxIndex, short button, short shift, long x, long y, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowIpaddrA(long hWnd);
	void __stdcall FieldTextChangedIpaddrA(long editBoxIndex);
	void __stdcall KeyDownIpaddrA(long editBoxIndex, short* keyCode, short shift);
	void __stdcall KeyPressIpaddrA(long editBoxIndex, short* keyAscii);
	void __stdcall KeyUpIpaddrA(long editBoxIndex, short* keyCode, short shift);
	void __stdcall MClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseDownIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseEnterIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseHoverIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseLeaveIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseMoveIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseUpIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall MouseWheelIpaddrA(long editBoxIndex, short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveIpaddrA(long editBoxIndex, LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveIpaddrA(long editBoxIndex, LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall RClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall RDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowIpaddrA(long hWnd);
	void __stdcall ResizedControlWindowIpaddrA();
	void __stdcall XClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);
	void __stdcall XDblClickIpaddrA(long editBoxIndex, short button, short shift, long x, long y);

	void __stdcall ClickHkboxA(short button, short shift, long x, long y);
	void __stdcall ContextMenuHkboxA(short button, short shift, long x, long y);
	void __stdcall DblClickHkboxA(short button, short shift, long x, long y);
	void __stdcall DestroyedControlWindowHkboxA(long hWnd);
	void __stdcall KeyDownHkboxA(short* keyCode, short shift);
	void __stdcall KeyPressHkboxA(short* keyAscii);
	void __stdcall KeyUpHkboxA(short* keyCode, short shift);
	void __stdcall MClickHkboxA(short button, short shift, long x, long y);
	void __stdcall MDblClickHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseDownHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseEnterHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseHoverHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseLeaveHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseMoveHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseUpHkboxA(short button, short shift, long x, long y);
	void __stdcall MouseWheelHkboxA(short button, short shift, long x, long y, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragEnterHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall OLEDragLeaveHkboxA(LPDISPATCH data, short button, short shift, long x, long y);
	void __stdcall OLEDragMouseMoveHkboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y);
	void __stdcall RClickHkboxA(short button, short shift, long x, long y);
	void __stdcall RDblClickHkboxA(short button, short shift, long x, long y);
	void __stdcall RecreatedControlWindowHkboxA(long hWnd);
	void __stdcall ResizedControlWindowHkboxA();
	void __stdcall XClickHkboxA(short button, short shift, long x, long y);
	void __stdcall XDblClickHkboxA(short button, short shift, long x, long y);

	void __stdcall BeforeDrawTextUdtxtboxA();
	void __stdcall ChangedAcceleratorsUdtxtboxA(LPDISPATCH accelerators);
	void __stdcall ChangingAcceleratorsUdtxtboxA(LPDISPATCH accelerators, VARIANT_BOOL* cancelChanges);
	void __stdcall ClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall ContextMenuUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails, VARIANT_BOOL* showDefaultMenu);
	void __stdcall DblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall DestroyedControlWindowUdtxtboxA(long hWnd);
	void __stdcall KeyDownUdtxtboxA(short* keyCode, short shift);
	void __stdcall KeyPressUdtxtboxA(short* keyAscii);
	void __stdcall KeyUpUdtxtboxA(short* keyCode, short shift);
	void __stdcall MClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseDownUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseEnterUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseHoverUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseLeaveUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseMoveUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseUpUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall MouseWheelUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails, long scrollAxis, short wheelDelta);
	void __stdcall OLEDragDropUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragEnterUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragLeaveUdtxtboxA(LPDISPATCH data, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OLEDragMouseMoveUdtxtboxA(LPDISPATCH data, EditCtlsLibA::OLEDropEffectConstants* effect, short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall OutOfMemoryUdtxtboxA();
	void __stdcall RClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall RecreatedControlWindowUdtxtboxA(long hWnd);
	void __stdcall ResizedControlWindowUdtxtboxA();
	void __stdcall TextChangedUdtxtboxA();
	void __stdcall TruncatedTextUdtxtboxA();
	void __stdcall ValueChangedUdtxtboxA();
	void __stdcall ValueChangingUdtxtboxA(long currentValue, long delta, VARIANT_BOOL* cancelChange);
	void __stdcall WritingDirectionChangedUdtxtboxA(EditCtlsLibA::WritingDirectionConstants newWritingDirection);
	void __stdcall XClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
	void __stdcall XDblClickUdtxtboxA(short button, short shift, long x, long y, EditCtlsLibA::HitTestConstants hitTestDetails);
};
