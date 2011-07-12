# PYTHON 2.7

import sys,os
import time
import thread
import atexit
from datetime import datetime
import Image


#from pyglet import app
#from pyglet import window

from ctypes import *
from ctypes.wintypes import  HANDLE, ULONG, DWORD, BOOL, LPCSTR, LPCWSTR, WinError, HRESULT, WORD, BYTE
from ctypes.util import find_library

import win32event
from win32process import beginthreadex, GetExitCodeProcess
from win32api import GetCurrentThreadId
from win32ui import CreateThread

from threading import Thread


"""
 version 0.01
 
 Please check Changes.txt for changes and the development planning
"""

# NUI Common Initialization Declarations
NUI_INITIALIZE_FLAG_USES_DEPTH_AND_PLAYER_INDEX         = 0x00000001
NUI_INITIALIZE_FLAG_USES_COLOR                          = 0x00000002
NUI_INITIALIZE_FLAG_USES_SKELETON                       = 0x00000008  
NUI_INITIALIZE_FLAG_USES_DEPTH                          = 0x00000020
NUI_INITIALIZE_DEFAULT_HARDWARE_THREAD                  = 0xFFFFFFFF

# _NUI_IMAGE_TYPE
NUI_IMAGE_TYPE_DEPTH_AND_PLAYER_INDEX                   = 0 #/ USHORT
NUI_IMAGE_TYPE_COLOR                                    = 1 # RGB32 data
NUI_IMAGE_TYPE_COLOR_YUV                                = 2 # YUY2 stream from camera h/w but converted to RGB32 before user getting it.
NUI_IMAGE_TYPE_COLOR_RAW_YUV                            = 3 # YUY2 stream from camera h/w.
NUI_IMAGE_TYPE_DEPTH                                    = 4 # USHORT
NUI_IMAGE_TYPE_DEPTH_AND_PLAYER_INDEX_IN_COLOR_SPACE    = 5 # not available yet
NUI_IMAGE_TYPE_DEPTH_IN_COLOR_SPACE                     = 6 # not available yet
NUI_IMAGE_TYPE_COLOR_IN_DEPTH_SPACE                     = 7

# _NUI_IMAGE_RESOLUTION
NUI_IMAGE_RESOLUTION_INVALID                            =-1
NUI_IMAGE_RESOLUTION_80x60                              = 0
NUI_IMAGE_RESOLUTION_320x240                            = 1
NUI_IMAGE_RESOLUTION_640x480                            = 2
NUI_IMAGE_RESOLUTION_1280x1024                          = 3 # for hires color only

NUI_IMAGE_PLAYER_INDEX_SHIFT                            = 3
NUI_IMAGE_PLAYER_INDEX_MASK                             = ((1 << NUI_IMAGE_PLAYER_INDEX_SHIFT)-1)
NUI_IMAGE_DEPTH_MAXIMUM                                 = ((4000 << NUI_IMAGE_PLAYER_INDEX_SHIFT) | NUI_IMAGE_PLAYER_INDEX_MASK)
NUI_IMAGE_DEPTH_MINIMUM                                 = (800 << NUI_IMAGE_PLAYER_INDEX_SHIFT)
NUI_IMAGE_DEPTH_NO_VALUE                                = 0

NUI_CAMERA_DEPTH_NOMINAL_FOCAL_LENGTH_IN_PIXELS         = 285.63   # Based on 320x240 pixel size.
NUI_CAMERA_DEPTH_NOMINAL_INVERSE_FOCAL_LENGTH_IN_PIXELS = 3.501e-3 # (1/NUI_CAMERA_DEPTH_NOMINAL_FOCAL_LENGTH_IN_PIXELS)
NUI_CAMERA_DEPTH_NOMINAL_DIAGONAL_FOV                   = 70.0
NUI_CAMERA_DEPTH_NOMINAL_HORIZONTAL_FOV                 = 58.5
NUI_CAMERA_DEPTH_NOMINAL_VERTICAL_FOV                   = 45.6

NUI_CAMERA_COLOR_NOMINAL_FOCAL_LENGTH_IN_PIXELS         = 531.15   # Based on 640x480 pixel size.
NUI_CAMERA_COLOR_NOMINAL_INVERSE_FOCAL_LENGTH_IN_PIXELS = 1.83e-3  # (1/NUI_CAMERA_COLOR_NOMINAL_FOCAL_LENGTH_IN_PIXELS)
NUI_CAMERA_COLOR_NOMINAL_DIAGONAL_FOV                   = 73.9
NUI_CAMERA_COLOR_NOMINAL_HORIZONTAL_FOV                 = 62.0
NUI_CAMERA_COLOR_NOMINAL_VERTICAL_FOV                   = 48.6

# the max # of NUI output frames you can hold w/o releasing
NUI_IMAGE_STREAM_FRAME_LIMIT_MAXIMUM                    = 4

# return S_FALSE instead of E_NUI_FRAME_NO_DATA if # NuiImageStreamGetNextFrame( )
# doesn't have a frame ready and a timeout != INFINITE is used
NUI_IMAGE_STREAM_FLAG_SUPPRESS_NO_FRAME_DATA            = 0x00010000

# Camera extreme angles
Camera_ElevationMaximum                                 = 25 # actual 27 till -27
Camera_ElevationMinimum                                 =-25

# others
S_OK                                                    = 0
#HANDLE = HRESULT
####################

# using ctypes to talk to kinect drivers
# http://docs.python.org/library/ctypes.html
# http://python.net/crew/theller/ctypes/tutorial.html#calling-functions


# create event handlers
m_hEvNuiProcessStop    = win32event.CreateEvent(None, 0, 0, None)
m_hNextDepthFrameEvent = win32event.CreateEvent(None, 1, 0, None)
m_hNextVideoFrameEvent = win32event.CreateEvent(None, 1, 0, None)
m_hNextSkeletonEvent   = win32event.CreateEvent(None, 1, 0, None)

hEvents = [ m_hEvNuiProcessStop,  
            m_hNextDepthFrameEvent,
            m_hNextVideoFrameEvent,
            m_hNextSkeletonEvent ]

#m_pVideoStreamHandle = pointer(c_int(1))
#m_pDepthStreamHandle = pointer(c_int(1))

m_pVideoStreamHandle = c_void_p()
m_pDepthStreamHandle = c_void_p()

class _NUI_IMAGE_VIEW_AREA(Structure):
    _fields_ = [("eDigitalZoom_NotUsed",c_int),
                 ("lCenterX_NotUsed",c_long),
                 ("lCenterY_NotUsed",c_long)]

class _NuiImageBuffer(Structure):
    _fields_= [("m_Width",c_int),
                ("m_Height",c_int),
                ("m_BytesPerPixel",c_int),
                ("m_pBuffer",c_void_p),
                ("BufferLen",c_int),        # m_Width * m_Height * m_BytesPerPixel;
                ("Pitch",c_int)]            # m_Width * m_BytesPerPixel;

class _NUI_IMAGE_FRAME(Structure):
    _fields_ = [("liTimeStamp", c_longlong),
                ("dwFrameNumber", c_ulong),
                ("eImageType", c_int),               # enum NUI_IMAGE_TYPE
                ("eResolution", c_int),              # enum NUI_IMAGE_RESOLUTION
                ("pFrameTexture", POINTER(_NuiImageBuffer)),  # obj from class NuiImageBuffer
                ("dwFrameFlags_NotUsed", c_ulong),  
                ("ViewArea_NotUsed", _NUI_IMAGE_VIEW_AREA)]

pImageFrame = _NUI_IMAGE_FRAME()

class KinectInterface:
    # dll files
    dll_Kinect  = "MSRKINECTNUI.DLL"
    lock        = thread.allocate_lock()
    nrKinects   = 0
    
    def __init__(self):
        if 'win32' != sys.platform:
            raise Exception('Only supports x86 windows.')
        
        if find_library(self.dll_Kinect):
            try: 
                self.dll = windll.LoadLibrary(self.dll_Kinect)
            except:
                raise Exception('Kinect drivers installed but can not be loaded')
        else:
            raise Exception('Kinect drivers not installed.')

        # self.m_hThNuiProcess = Nui_ProcessThread("m_hThNuiProcess", self)
        
        # check how many kinects are connected to the system
        self.nrKinects = self.nrKinectsConnected()
        
        atexit.register(self.goodbye)
        
        if self.nrKinects == 0:
            raise Exception('No kinect connected')
        

    def Nui_ProcessThread(self,*args):
        print "Nui_ProcessThread started:",args
        # http://stackoverflow.com/questions/100624/python-on-windows-how-to-wait-for-multiple-child-processes

        abort=0
        while not abort:
            nEventIdx = win32event.WaitForMultipleObjects(hEvents, False, 100)
            if nEventIdx == 0:
                print "m_hEvNuiProcessStop event occurred:",hEvents[nEventIdx]
                abort = 1
            elif nEventIdx == 1:
                print "m_hNextDepthFrameEvent event occurred:",hEvents[nEventIdx]
            elif nEventIdx == 2:
                print "m_hNextVideoFrameEvent event occurred:",hEvents[nEventIdx]
                self.Nui_GotVideoAlert()
            elif nEventIdx == 3:
                print "m_hNextSkeletonEvent event occurred:",hEvents[nEventIdx]
                self.Nui_GotVideoAlert()
            else:
                print "unknown nEventIdx:",nEventIdx
                    
            #finished    = handles[nEventIdx]
            #exitcode    = GetExitCodeProcess(finished)
            #procname    = hEvents.pop(finished)
            #finished.close()
            #print "Subprocess %s finished with exit code %d" % (procname, exitcode)
            time.sleep(0.3)
            
        print "Nui_ProcessThread ended"

    def getImageFrame(self):
        hr=cdll.MSRKINECTNUI.NuiImageGetNextFrame()
        if hr != S_OK:
            print "getImageFrame failed:",hr

    def getSkeletonFrame(self):
        hr=cdll.MSRKINECTNUI.NuiSkeletonGetNextFrame()
        if hr != S_OK:
            print "getSkeletonFrame failed:",hr
    
    def Nui_Init(self):
        # todo: use pyglet to handle video window        
        self.NuiInitialize()
        #self.NuiSkeletonTrackingEnable()
        self.NuiImageVideoStreamOpen()
        #self.NuiImageDepthStreamOpen()
               
        # create the worker threads
        ProcessThread=Thread(target=self.Nui_ProcessThread, args=())
        ProcessThread.setDaemon(True)
        ProcessThread.start()
        
        while ProcessThread.isAlive():
            time.sleep(1)
    
    def NuiInitialize(self,options = None):
        if options == None:
            options = NUI_INITIALIZE_FLAG_USES_COLOR | \
                      NUI_INITIALIZE_FLAG_USES_DEPTH_AND_PLAYER_INDEX | \
                      NUI_INITIALIZE_FLAG_USES_SKELETON
                    
        hr=cdll.MSRKINECTNUI.NuiInitialize( options )
        if hr != S_OK:
            raise Exception('NuiInitialize failed.')
        

    def NuiSkeletonTrackingEnable(self):
        print "NuiSkeletonTrackingEnable"
        cdll.MSRKINECTNUI.NuiSkeletonTrackingEnable.restype= c_uint
        cdll.MSRKINECTNUI.NuiSkeletonTrackingEnable.argtypes = [HRESULT, c_uint]
        
        hr = cdll.MSRKINECTNUI.NuiSkeletonTrackingEnable( m_hNextSkeletonEvent, c_uint(0))
        if hr != S_OK:
            raise Exception('NuiSkeletonTrackingEnable failed.')

    def NuiImageVideoStreamOpen(self):
        print "NuiImageVideoStreamOpen"
        cdll.MSRKINECTNUI.NuiImageStreamOpen.argtypes = [c_int, c_int,c_int,c_int, HRESULT , c_void_p]
        cdll.MSRKINECTNUI.NuiImageStreamOpen.restype  = c_uint
        hr = cdll.MSRKINECTNUI.NuiImageStreamOpen(NUI_IMAGE_TYPE_COLOR,
                                                  NUI_IMAGE_RESOLUTION_640x480,
                                                  0,2,
                                                  m_hNextVideoFrameEvent,
                                                  byref(m_pVideoStreamHandle))
        if hr != S_OK:
            raise Exception('NuiImageStreamOpen (NuiImageVideoStreamOpen) failed.')

    def NuiImageDepthStreamOpen(self):
        print "NuiImageDepthStreamOpen"
        cdll.MSRKINECTNUI.NuiImageStreamOpen.argtypes = [c_uint, c_uint,c_uint,c_uint, HRESULT , c_void_p]
        cdll.MSRKINECTNUI.NuiImageStreamOpen.restype  = c_uint
        hr = cdll.MSRKINECTNUI.NuiImageStreamOpen(NUI_IMAGE_TYPE_DEPTH_AND_PLAYER_INDEX,
                                                  NUI_IMAGE_RESOLUTION_320x240,
                                                  0,2,
                                                  m_hNextDepthFrameEvent,
                                                  byref(_pDepthStreamHandle))
        if hr != S_OK:
            raise Exception('NuiImageStreamOpen (NuiImageDepthStreamOpen) failed.')

    def nrKinectsConnected(self):
        nrKinects = c_int(0)
        hr=cdll.MSRKINECTNUI.MSR_NUIGetDeviceCount(byref(nrKinects))
        if hr != S_OK:
            raise Exception('error white getting nr of Kinects.')
        print "Kinects connected:",nrKinects.value
        return nrKinects.value
    
    def Nui_GotVideoAlert(self):
        pImageFrame.values = None
        
        # reserve buffer
        #pImageFrame.pFrameTexture = c_ubyte * (640*480)
        #pImageFrame.pFrameTexture = pImageFrame.pFrameTexture()
        
        cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame.argtypes = [c_void_p, c_uint, POINTER(_NUI_IMAGE_FRAME)]
        cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame.restype  = c_uint
        hr=cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame(m_pVideoStreamHandle,0,byref(pImageFrame));
        
        #cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame.argtypes = [c_void_p, c_uint, _NUI_IMAGE_FRAME]
        #cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame.restype  = c_uint
        #hr=cdll.MSRKINECTNUI.NuiImageStreamGetNextFrame(m_pVideoStreamHandle,0,pImageFrame);
            
        if hr != S_OK:
            raise Exception('Nui_GotVideoAlert failed.')
        
        
        print "timeStamp:",pImageFrame.liTimeStamp,datetime.fromtimestamp(pImageFrame.liTimeStamp).strftime("%d/%m/%y %H:%M")
        
        #pImage = pImageFrame.pFrameTexture
        #INTP = POINTER(c_long)
        #print "pImage:",type(pImage)
        #addr = addressof(pImage)
        #print ' address: %x'%(addr), type(addr)
        #ptr = cast(addr, INTP)
        #print ' pointer:', ptr
        #print ' value:', ptr[0]

        #graydata = cast(graydata, POINTER(c_ubyte))

        #self.convertImageFromString(imgSize,buffer)

        
        #mgBuffer=pImage.values
        #imgBuffer = c_ubyte * (640*480)()
        #cast(imgBuffer, POINTER(c_int))
        
        #cast(pImage,)
        
        #imgBuffer = cast(pImage, POINTER(c_ubyte))
        
        #print "imgBuffer:",type(imgBuffer)

        #imgSize = (int(640), int(480))
        #self.convertImage(imgSize,pImage)
        
        cdll.MSRKINECTNUI.NuiImageStreamReleaseFrame( m_pVideoStreamHandle, pImageFrame );

        # KINECT_LOCKED_RECT LockedRect;
        
        # pTexture->LockRect( 0, &LockedRect, NULL, 0 );
        
        # if( LockedRect.Pitch != 0 )
        # {
        #     BYTE * pBuffer = (BYTE*) LockedRect.pBits;    
        #     m_DrawVideo.DrawFrame( (BYTE*) pBuffer );
        # }
        # else
        # {
        #     OutputDebugString( L"Buffer length of received texture is bogus\r\n" );
        # }
    
    def convertImageFromString(self,imgSize,imgString):
        im = Image.fromstring('L', imgSize, imgString)
        im.save('out.png')
    
    def convertImage(self,imgSize,bytes):
        im = Image.frombuffer('L', (10,10), bytes,'raw', 'L', 0, 1)
        #im = Image.fromarray('L', bytes)
        #im = Image.fromstring('L', imgSize, bytes) #, 'raw', 'F;16')
        im.save('out.png')
        
        #PlanarImage Image = e.ImageFrame.Image;
        #video.Source = BitmapSource.Create(
        #Image.Width, Image.Height, 96, 96, PixelFormats.Bgr32, null, 
        #Image.Bits, Image.Width * Image.BytesPerPixel)
  
    def __del__(self):
        print "Shutting down kinect"
        cdll.MSRKINECTNUI.NuiShutdown()
        #windll.kernel32.CloseHandle(lpProcessInformation.hProcess)

    def goodbye(self):
        print "Application ended."

    def getAngle(self):
        angle = c_int(0)
        cdll.MSRKINECTNUI.NuiCameraElevationGetAngle(byref(angle))
        return angle.value
        
    def setAngle(self,angle=0):
        print "moving Kinect to angle",angle
        if self.getAngle() == angle:
            return
        angle=max(angle,Camera_ElevationMinimum) # clamp the angle
        angle=min(angle,Camera_ElevationMaximum) # bewteen min and max
        hr = cdll.MSRKINECTNUI.NuiCameraElevationSetAngle(c_int(angle));
        pa = self.getAngle()
        time.sleep(1)
        while (self.getAngle() != pa):
            time.sleep(0.1)
            pa = self.getAngle()
            
    def testAll(self):
        print("== Test motor")
        print("  moving to 0 degrees")
        myKinect.setAngle(0)
        print("  moving to -20 degrees")
        myKinect.setAngle(-20)
        print("  moving to  20 degrees")
        myKinect.setAngle(20)
        print("  moving to 0 degrees")
        myKinect.setAngle(0)
        print("== Test motor ended")


#class pygletApp(window.Window):
#    abort=0
#    def __init__(self):
#        super(pygletApp, self).__init__(resizable=True,width=640,height=480,vsync=False)
#    
#    def on_key_press(self,sym, mod):
#        if sym == key.ESCAPE:
#            # stop our program
#            self.has_exit = True
#            
#    def update(self):
#        pass
#        
#    def draw(self):
#        pass


## pyglet stuff
#window = window.Window()

#@window.event
#def on_draw():
#    window.clear()
#    print "d",


# simple unit test
if __name__ == '__main__':
    myKinect = KinectInterface()

    if myKinect.nrKinects == 0:
        print("exiting.")
        os._exit(-1)
        
        
    myKinect.Nui_Init()
    #...

    del myKinect
    os._exit(0)

