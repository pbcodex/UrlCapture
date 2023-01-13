; Author RASHAD  & Kiffi (2013)

UsePNGImageEncoder()

Global sURL.s,imgn

Enumeration
  #WinSnapShot
  #WebSnapShot
EndEnumeration

#DVASPECT_CONTENT  = 1
#DVASPECT_THUMBNAIL = 2
#DVASPECT_ICON = 4
#DVASPECT_DOCPRINT = 8


Procedure CaptureWebPage(URL$, WBWidth.i, WBHeight.i, Filename.s)
  ;----------------------------------------------------------------
  
  Define.IWebBrowser2   m_pWebBrowser
  
  Define.IHTMLDocument2 pDocument
  Define.IHTMLDocument3 pDocument3
  Define.IHTMLElement   pElement
  Define.IHTMLElement2  pElement2
  Define.iDispatch      pDispatch
  Define.IViewObject2   pViewObject
  
  Define.i bodyHeight
  Define.i bodyWidth
  Define.i rootHeight
  Define.i rootWidth
  
  Define.RECT rcBounds
  
  Define.i bolFlag
  Define.i IsBusy
  Define.i hr
  
  bolFlag = #False
  
  If OpenWindow(#WinSnapShot, 0, 0, 0, 0, "", #PB_Window_Invisible | #PB_Window_BorderLess)
    
    WebGadget(#WebSnapShot, 0, 0, 0, 0, URL$)
    
    Repeat
      If GetGadgetAttribute(#WebSnapShot, #PB_Web_Busy) = 0
        Break
      EndIf
      
      While WindowEvent(): Delay(1) : Wend
    ForEver
    
    m_pWebBrowser = GetWindowLongPtr_(GadgetID(#WebSnapShot), #GWL_USERDATA)
    
    hr = m_pWebBrowser\get_document(@pDispatch)
    If hr = #S_OK
      
      If pDispatch
        
        hr = pDispatch\QueryInterface(?IID_IHTMLDocument2, @pDocument)
        If hr = #S_OK
          
          If pDocument
            
            hr = pDocument\get_body(@pElement)
            If hr = #S_OK
              If pElement
                
                hr = pElement\QueryInterface(?IID_IHTMLElement2, @pElement2)
                If hr = #S_OK
                  If pElement2
                    
                    hr = pElement2\get_scrollHeight(@bodyHeight)
                    If hr = #S_OK
                      ;Debug "bodyHeight: " + Str(bodyHeight)
                      
                      hr = pElement2\get_scrollWidth(@bodyWidth)
                      If hr = #S_OK
                        ;Debug "bodyWidth: " + Str(bodyWidth)
                        
                        hr = pDispatch\QueryInterface(?IID_IHTMLDocument3, @pDocument3)
                        If hr = #S_OK
                          
                          If pDocument3
                            
                            hr = pDocument3\get_documentElement(@pElement)
                            If hr <> #S_OK : ProcedureReturn #False : EndIf
                            
                            hr = pElement\QueryInterface(?IID_IHTMLElement2, @pElement2)
                            If hr <> #S_OK : ProcedureReturn #False : EndIf
                            
                            hr = pElement2\get_scrollHeight(@rootHeight)
                            If hr <> #S_OK : ProcedureReturn #False : EndIf
                            ;Debug "rootHeight: " + Str(rootHeight)
                            
                            hr = pElement2\get_scrollWidth(@rootWidth)
                            If hr <> #S_OK : ProcedureReturn #False : EndIf
                            ;Debug "rootWidth: " + Str(rootWidth)
                            
                            Define.i width
                            Define.i height
                            width = bodyWidth
                            If rootWidth > bodyWidth : width = rootWidth : EndIf
                            
                            height = bodyHeight
                            If rootHeight > bodyHeight : height = rootHeight : EndIf
                            
                            width + 22
                            
                            ResizeGadget(#WebSnapShot, 0, 0, width, height)
                            
                            hr = m_pWebBrowser\QueryInterface(?IID_IViewObject2, @pViewObject)
                            
                            If hr = #S_OK
                              
                              If pViewObject
                                
                                Define.i hdcMain
                                
                                hdcMain = GetDC_(0)
                                If hdcMain
                                  
                                  Define.i hdcMem
                                  
                                  hdcMem  = CreateCompatibleDC_(hdcMain)
                                  If hdcMem
                                    
                                    Define.i hBitmap
                                    
                                    hBitmap = CreateCompatibleBitmap_(hdcMain, width, height)
                                    If hBitmap
                                      
                                      Define.i oldImage
                                      
                                      oldImage = SelectObject_(hdcMem, hBitmap)
                                      
                                      rcBounds\top = 0
                                      rcBounds\left = 0
                                      rcBounds\right = width
                                      rcBounds\bottom = height
                                      
                                      pViewObject\Draw(#DVASPECT_CONTENT, -1, 0, 0, hdcMain, hdcMem, rcBounds, 0, 0, 0)
                                      
                                      Define.i Image
                                      
                                      Image = CreateImage(#PB_Any, width, height)
                                      If Image
                                        
                                        Define.i img_hDC
                                        
                                        img_hDC = StartDrawing(ImageOutput(Image))
                                        If img_hDC
                                          
                                          BitBlt_(img_hDC, 0, 0, width, height, hdcMem, 0, 0, #SRCCOPY)
                                          StopDrawing()
                                          
                                          SaveImage(Image,Filename,#PB_ImagePlugin_PNG)
                                          bolFlag = #True
                                          
                                        EndIf ; img_hDC
                                        
                                        FreeImage(Image)
                                        
                                      EndIf ; Image
                                      
                                      SelectObject_(hdcMem, oldImage)
                                      
                                    EndIf ; hBitmap
                                    
                                    DeleteDC_(hdcMem)
                                    
                                  EndIf ; hdcMem
                                  
                                  ReleaseDC_(0, hdcMain)
                                  
                                EndIf ; hdcMain
                                
                                pViewObject\Release()
                                
                              EndIf ; pViewObject
                            EndIf   ; HR
                            
                            pDocument3\Release()
                            
                          EndIf ; pDocument3
                          
                        EndIf ; HR
                        
                      EndIf ; HR
                      
                    EndIf ; HR
                    
                    pElement2\Release()
                    
                  EndIf ; pElement2
                  
                EndIf ; HR
                
                pElement\Release()
                
              EndIf ; pElement
              
            EndIf ; HR
            
            pDocument\Release()
            
          EndIf ; pDocument
          
        EndIf ; HR
        
        pDispatch\Release()
        
      EndIf ; pDispatch
      
    EndIf ; HR
    
    CloseWindow(#WinSnapShot)
    
  EndIf
  
  ProcedureReturn bolFlag
  
EndProcedure

Procedure Snapshot()
  ;-------------------
  
  ;Define sURL.s
  Define sFileOut.s
  Define sNow.s
  
  
  ExamineDesktops()
  sFileOut = GetUserDirectory(#PB_Directory_Downloads)+"web_image_"+Str(imgn)+".png"
  imgn+1
  If CaptureWebPage(sURL, DesktopWidth(0), DesktopHeight(0), sFileOut)
    
  Else
    MessageRequester("Web Page", "Capture Failed")
  EndIf
  
EndProcedure

;Snapshot()
LoadFont(0,"Broadway",24)

OpenWindow(10,0,0,400,80,"Capture Web",#PB_Window_SystemMenu |#PB_Window_ScreenCentered)
StickyWindow(10,1)
StringGadget(20,10,10,380,24,"")
ButtonGadget(30,10,50,60,24,"Capture")
TextGadget(40,140,50,80,40,"")
SetGadgetFont(40,0)
SetGadgetColor(40,#PB_Gadget_FrontColor,#Red)


Repeat
  Select WaitWindowEvent()
      
    Case #PB_Event_CloseWindow
      Quit = 1
      
    Case #PB_Event_Gadget
      Select EventGadget()
        Case 30
          SetGadgetText(40,"")
          sURL.s = GetGadgetText(20)
          If sURL.s
            Snapshot()
          EndIf
          SetGadgetText(20,"")
          SetGadgetText(40,"FINISHED")
          
      EndSelect          
      
  EndSelect 
  
Until Quit = 1
End

End


DataSection
  
  IID_IHTMLDocument2:
  ;332C4425-26CB-11D0-B483-00C04FD90119
  Data.i $332C4425
  Data.w $26CB, $11D0
  Data.b $B4, $83, $00, $C0, $4F, $D9, $01, $19
  
  IID_IHTMLDocument3:
  ;3050F485-98B5-11CF-BB82-00AA00BDCE0B
  Data.i $3050F485
  Data.w $98B5, $11CF
  Data.b $BB, $82, $00, $AA, $00, $BD, $CE, $0B
  
  IID_IHTMLElement2:
  ;3050f434-98b5-11cf-bb82-00aa00bdce0b
  Data.i $3050F434
  Data.w $98B5, $11CF
  Data.b $BB, $82, $00, $AA, $00, $BD, $CE, $0B
  
  IID_IViewObject2:
  ;00000127-0000-0000-c000-000000000046
  Data.i $00000127
  Data.w $0000, $0000
  Data.b $C0, $00, $00, $00, $00, $00, $00, $46
  
  
EndDataSection
