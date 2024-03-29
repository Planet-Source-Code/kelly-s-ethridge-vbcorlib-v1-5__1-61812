Attribute VB_Name = "modHelper"
'    CopyRight (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modHelper
'

''
' Creates an object that provides ASM code for special functions.
'
Option Explicit

Public Helper As Helper
Public Ecvt As Ecvt

Private Type HelperType
    pVTable As Long
    Func(12) As Long
End Type

Private mHelper As HelperType
Private mAsm() As Long
Private mMSVCLib As Long


''
' Returns an ASM memory swapping routine.
'
' @param Size The number of bytes to be swapped.
' @return A swapper that swaps the specified number of bytes at a time.
'
Public Function GetSwapper(ByVal Size As Long) As ISwap
    Select Case Size
        Case 4:     Set GetSwapper = NewDelegator(mHelper.Func(3))
        Case 8:     Set GetSwapper = NewDelegator(mHelper.Func(4))
        Case 16:    Set GetSwapper = NewDelegator(mHelper.Func(5))
        Case 2:     Set GetSwapper = NewDelegator(mHelper.Func(6))
        Case 1:     Set GetSwapper = NewDelegator(mHelper.Func(7))
        Case Else
            Throw Cor.NewArgumentOutOfRangeException("Not a valid swapper size. Must be {1,2,4,8,16}.")
    End Select
End Function

''
' Creates the helper object.
'
Public Sub InitHelper()
    Dim this As Long
    InitAsm
    this = CoTaskMemAlloc(LenB(mHelper))
    If this = 0 Then Err.Raise 7
    
    With mHelper
        .Func(0) = FuncAddr(AddressOf QueryInterface)
        .Func(1) = FuncAddr(AddressOf AddRefRelease)
        .Func(2) = .Func(1)
        .Func(3) = VarPtr(mAsm(0))
        .Func(4) = VarPtr(mAsm(5))
        .Func(5) = VarPtr(mAsm(13))
        .Func(6) = VarPtr(mAsm(25))
        .Func(7) = VarPtr(mAsm(31))
        .Func(8) = VarPtr(mAsm(36))
        .Func(9) = VarPtr(mAsm(39))
        .Func(10) = VarPtr(mAsm(50))
        .Func(11) = VarPtr(mAsm(72))
        .Func(12) = VarPtr(mAsm(76))
        
        .pVTable = this + 4
    End With
    
    Call CopyMemory(ByVal this, mHelper, LenB(mHelper))
    
    ObjectPtr(Helper) = this
End Sub

Private Sub InitAsm()
    ReDim mAsm(79)
    ' Swap4  from Matt Curland
    mAsm(0) = &H824448B
    mAsm(1) = &HC24548B
    mAsm(2) = &HA87088B
    mAsm(3) = &HCC20889
    mAsm(4) = &H90909000
        
    ' Swap8
    mAsm(5) = &H824448B
    mAsm(6) = &HC24548B
    mAsm(7) = &HA87088B
    mAsm(8) = &H488B0889
    mAsm(9) = &H44A8704
    mAsm(10) = &HC2044889
    mAsm(11) = &H9090000C
    mAsm(12) = &H90909090
    
    ' Swap16
    mAsm(13) = &H824448B
    mAsm(14) = &HC24548B
    mAsm(15) = &HA87088B
    mAsm(16) = &H488B0889
    mAsm(17) = &H44A8704
    mAsm(18) = &H8B044889
    mAsm(19) = &H4A870848
    mAsm(20) = &H8488908
    mAsm(21) = &H870C488B
    mAsm(22) = &H48890C4A
    mAsm(23) = &HCC20C
    mAsm(24) = &H33909090
        
    ' Swap2
    mAsm(25) = &H824448B
    mAsm(26) = &HC24548B
    mAsm(27) = &H66088B66
    mAsm(28) = &H89660A87
    mAsm(29) = &HCC208
    mAsm(30) = &H33909090

    ' Swap1
    mAsm(31) = &H824448B
    mAsm(32) = &HC24548B
    mAsm(33) = &HA86088A
    mAsm(34) = &HCC20888
    mAsm(35) = &H90909000

    ' DerefEBP  from Matt Curland
    mAsm(36) = &H8244C8B
    mAsm(37) = &HD448B
    mAsm(38) = &H900008C2

    ' MoveVariant from Matt Curland
    mAsm(39) = &HC24448B
    mAsm(40) = &H824548B
    mAsm(41) = &H8B56C88B
    mAsm(42) = &H8B328931
    mAsm(43) = &H72890471
    mAsm(44) = &H8718B04
    mAsm(45) = &H5E087289
    mAsm(46) = &H890C498B
    mAsm(47) = &HC7660C4A
    mAsm(48) = &HC2000000
    mAsm(49) = &H9090000C

    ' _ecvt call
    mAsm(50) = &H81EC8B55
    mAsm(51) = &HC0EC&
    mAsm(52) = &H57565300
    mAsm(53) = &HFF40BD8D
    mAsm(54) = &H30B9FFFF
    mAsm(55) = &HB8000000
    mAsm(56) = &HCCCCCCCC
    mAsm(57) = &H458BABF3
    mAsm(58) = &H4D8B501C
    mAsm(59) = &H558B5118
    mAsm(60) = &H45DD5214
    mAsm(61) = &H8EC830C
    mAsm(62) = &HB8241CDD
    mAsm(63) = &HFFFFF3EC   ' ecvt address goes here
    mAsm(64) = &H9090D0FF
    mAsm(65) = &H5F14C483
    mAsm(66) = &HC4815B5E
    mAsm(67) = &HC0&
    mAsm(68) = &H9090EC3B
    mAsm(69) = &H8B909090
    mAsm(70) = &H18C25DE5
    mAsm(71) = &H90909000
    
    ' compatible libraries
    ' msvcrt20.dll
    ' msvcrt40.dll
    ' msvcr70.dll
    ' msvcr71.dll
    ' msvcr71d.dll
    mMSVCLib = LoadLibrary("msvcrt.dll")
    mAsm(63) = GetProcAddress(mMSVCLib, "_ecvt")
    
    'shift right
    mAsm(72) = &H824448B
    mAsm(73) = &HC244C8B
    mAsm(74) = &HCC2E8D3
    mAsm(75) = &HCCCCCC00
    
    'shift left
    mAsm(76) = &H824448B
    mAsm(77) = &HC244C8B
    mAsm(78) = &HCC2E0D3
    mAsm(79) = &HCCCCCC00
    
End Sub



Private Function QueryInterface(ByVal this As Long, ByVal riid As Long, pvObj As Long) As Long
    QueryInterface = E_NOINTERFACE
End Function
Private Function AddRefRelease(ByVal this As Long) As Long
    ' do nothing
    CoTaskMemFree this
End Function
