package main

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	//	"log"
	//	"strconv"
	"github.com/AllenDang/w32"
	"strings"
	"time"
)

func run_file(msaccess_file string, run_sub string, arg string) error {

	//fmt.Printf("%s",reflect.TypeOf(access))
	//

	start := time.Now()
	global.mutex.Lock()
	global.mutex_lock_started = time.Now()
	args := strings.Fields(arg)

	ole.CoInitialize(0)


	clsid, _ := ole.CLSIDFromString("{73A4C9C1-D68D-11D0-98BF-00A0C90DC8D9}")
	unknown, _ := ole.CreateInstance(clsid, ole.IID_IUnknown)
	access, _ := unknown.QueryInterface(ole.IID_IDispatch)
	


	r, err := oleutil.CallMethod(access, "OpenCurrentDatabase", msaccess_file, true)
	if err != nil {
		global.nats_c.Publish("log", "g305transport "+err.Error())
		w32.PostMessage(w32.HWND(oleutil.MustCallMethod(access, "hWndAccessApp").Val), w32.WM_QUIT, 0, 0)
		time.Sleep(time.Duration(1) * time.Second)
		oleutil.CallMethod(access, "Quit", 2)
		access.Release()
		panic(err.Error())
	}

	r.ToIDispatch()


	switch len(args) {
	case 0:
		oleutil.MustCallMethod(access, "Run", run_sub).ToIDispatch()
	case 1:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0]).ToIDispatch()
	case 2:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0], args[1]).ToIDispatch()
	case 3:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0], args[1], args[2]).ToIDispatch()
	case 4:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0], args[1], args[2], args[3]).ToIDispatch()
	case 5:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0], args[1], args[2], args[3], args[4]).ToIDispatch()
	case 6:
		oleutil.MustCallMethod(access, "Run", run_sub, args[0], args[1], args[2], args[3], args[4], args[5]).ToIDispatch()

	}

	oleutil.CallMethod(access, "CloseCurrentDatabase")
	

	elapsed := time.Since(start)

	//w32 was needed because somehow the "local system" account won't quit the application.
	w32.PostMessage(w32.HWND(oleutil.MustCallMethod(access, "hWndAccessApp").Val), w32.WM_QUIT, 0, 0)
	// sleep was mentioned on the net to give PostMessage time
	time.Sleep(time.Duration(1) * time.Second)
	oleutil.CallMethod(access, "Quit", 2)

	access.Release()
	//	time.Sleep(2000000000)
	//	ole.CoUninitialize()
	global.mutex.Unlock()
	fmt.Printf("\nvia OLE took %s\n", elapsed)

	return nil

}
