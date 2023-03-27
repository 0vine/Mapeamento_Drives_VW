#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Rodrigo Chaves

 Script Function:
	Autoit Object
	08/2022

#ce ----------------------------------------------------------------------------

#include-once

#include <AutoitObject.au3>

; Initialize AutoItObject
_AutoItObject_StartUp()

Func newObject()
	Local $oThis = _AutoItObject_Create()
	_AutoItObject_AddMethod($oThis, "AddProperty", "_Object_AddProperty")
	_AutoItObject_AddMethod($oThis, "AddMethod", "_Object_AddMethod")
	_AutoItObject_AddMethod($oThis, "Create", "_Object_Create")
	_AutoItObject_AddProperty($oThis, "Object", $ELSCOPE_PUBLIC, 0)
	_AutoItObject_AddProperty($oThis, "Parent", $ELSCOPE_PUBLIC, 0)
	Return $oThis
EndFunc

Func addObjectProperty($obj, $sByRefProperty, $sByRefContent = 0)
	Return _AutoItObject_AddProperty($obj, $sByRefProperty, $ELSCOPE_PUBLIC, $sByRefContent)
EndFunc

Func addObjectMethod($obj, $sByRefMethod, $sByRefContent = 0)
	Return _AutoItObject_AddMethod($obj, $sByRefMethod, $sByRefContent)
EndFunc