' ExpandAnimations (https://github.com/monperrus/ExpandAnimations)

' Copyright 2009-2011 Matthew Neeley.
' Copyright 2011 Martin Monperrus.

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.

' You should have received a copy of the GNU Lesser General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

Private ANIMSET as String
Private ENUMACCESS as String
Private VISATTR as String


' expands the current document and saves it to PDF
' the expanded version is saved to disk in filename-expanded.odp
' the PDF is filename-expanded.pdf
sub Main 
  Dim doc As Object
  doc = thisComponent
  
  ' the expansion 
  newUrlPdf = expandAnimations(doc)
  
  msgbox "Expansion done! See "+newUrlPdf
end sub

' tests the module on /tmp/test.odp
' can be called on the command line with
' $ libreoffice "macro:///ExpandAnimations.ExpandAnimations.test"
sub test
  Dim Dummy()
  Url = "file:///home/martin/test-ExpandAnimations.odp"
  StarDesktop.loadComponentFromURL(Url, "_default", 0, Dummy)
  Main
end sub

' expands the animations and exports to PDF
function expandAnimations(doc as Object)

  ' rename the document
  docExpanded = renameAsExpanded(doc)
  
  ' expand it
  expandDocument(docExpanded)
  
  ' export to PDF
  ' returns the PDF file name
  expandAnimations= exportToPDF(docExpanded)  
  
  ' closing the expanded version  
  docExpanded.close(false)
  
end function


' renames the current document
' e.g. test.odp -> test-expanded.odp
function renameAsExpanded(doc as Object)
  Dim Dummy()
  
  If Not BasicLibraries.isLibraryLoaded("Tools") then
     BasicLibraries.loadLibrary("Tools")
  Endif

  sDocUrl = doc.getURL()
  sDocPath = DirectoryNameoutofPath(sDocUrl, "/")
  sDocFileName = FileNameoutofPath(sDocUrl, "/")
  sDocFileNameExtension = GetFileNameExtension(sDocUrl)
  sDocFileNameWithoutExtension = GetFileNameWithoutExtension(sDocUrl, "/") 
  
  newUrlExpanded = sDocPath + "/" + sDocFileNameWithoutExtension + "-expanded.odp"
  doc.storeToUrl(newUrlExpanded, Array())
  
  ' reloading the old document
  expandedDoc = StarDesktop.loadComponentFromURL(newUrlExpanded, "_default", 0, Dummy)  
  
  renameAsExpanded = expandedDoc
end function


' exports to PDF
function exportToPDF(doc as Object)
  If Not BasicLibraries.isLibraryLoaded("Tools") then
     BasicLibraries.loadLibrary("Tools")
  Endif

  sDocUrl = doc.getURL()
  sDocPath = DirectoryNameoutofPath(sDocUrl, "/")
  sDocFileName = FileNameoutofPath(sDocUrl, "/")
  sDocFileNameExtension = GetFileNameExtension(sDocUrl)
  sDocFileNameWithoutExtension = GetFileNameWithoutExtension(sDocUrl, "/") 

  newUrlPdf = sDocPath + "/" + sDocFileNameWithoutExtension + ".pdf"
  ' we use storeToUrl because we don't want to load the PDF
  ' actually we have to, otherwise, there is an error
  doc.storeToUrl(newUrlPdf, Array(makePropertyValue("FilterName", "impress_pdf_Export")))
  exportToPDF = newUrlPdf
end function 



' creates and returns a new com.sun.star.beans.PropertyValue.
' see http://www.oooforum.org/forum/viewtopic.phtml?t=5108
Function makePropertyValue( Optional cName As String, Optional uValue ) As com.sun.star.beans.PropertyValue
   Dim oPropertyValue As New com.sun.star.beans.PropertyValue
   If Not IsMissing( cName ) Then
      oPropertyValue.Name = cName
   EndIf
   If Not IsMissing( uValue ) Then
      oPropertyValue.Value = uValue
   EndIf
   makePropertyValue() = oPropertyValue
End Function 


function expandDocument(doc as Object)
    ANIMSET = "com.sun.star.animations.XAnimateSet"
    ENUMACCESS = "com.sun.star.container.XEnumerationAccess"
    VISATTR = "Visibility"
    
    numSlides = doc.getDrawPages().getCount()
        
    ' go through pages in reverse order
    for i = numSlides-1 to 0 step -1
        slide = doc.drawPages(i)
       	agordiGrandoDeTekstoj(slide)
        if hasAnimation(slide) then
            n = countAnimationSteps(slide)
            if n > 1 then
                origName = slide.Name
                replicateSlide(doc, slide, n)
                visArray = getShapeVisibility(slide, n)
                for frame = 0 to n-1
                    currentSlide = doc.drawPages(i + frame)
                    currentSlide.Name = origName & " (" & CStr(frame+1) & ")"
                    removeInvisibleShapes(currentSlide, visArray, frame)
                next
            end if
        end if
    next
    finalSlides = doc.getDrawPages().getCount()
    'MsgBox("Done! Expanded " & CStr(numSlides) & " slides to " & CStr(finalSlides) & ".")
    
    ' saving the expanded version
  doc.store()

end function


' Agordi grando de tekstoj (The English language discriminate by country of birth)
function agordiGrandoDeTekstoj(slide as Object)
  Dim numObjektoj As Integer
  Dim i As Integer
  Dim objekto As Object
  Dim eNum As Object
  Dim oTipo As String
  Dim oTeksto As Object
  Dim oAlghustigo As Integer
  Dim io(5) As Integer
  
  numObjektoj = slide.getCount()
  for i = 0 to numObjektoj-1
    objekto = slide.getByIndex(i)
    oTipo = objekto.GetShapeType()
    if oTipo = "com.sun.star.drawing.CustomShape" then
      objekto.TextAutoGrowWidth = false
    end if
    if oTipo = "com.sun.star.drawing.TextShape" then
      oAlghustigo = -1
      for kk = 0 to 4
        io(kk) = 0
      next
      eNum = objekto.Text.CreateEnumeration
      while eNum.HasMoreElements
        oTeksto = eNum.NextElement
        ' Ĉu la ero estas alineo?
        if oTeksto.SupportsService("com.sun.star.text.Paragraph") then
          if oTeksto.ParaAdjust < 4 then
            io(oTeksto.ParaAdjust) = 1
          end if
          select case oTeksto.ParaAdjust
            case 0
          	  if oAlghustigo = -1 then
          	    oAlghustigo = 0
          	  end if
          	case 1
          	  oAlghustigo = 2
          	case 2,4
          	  if oAlghustigo < 1 then
          	    oAlghustigo = 3
          	  end if
          	case 3
          	  if oAlghustigo < 1 then
          	    oAlghustigo = 1
          	  end if
          end select
        end if
      wend
      objekto.TextAutoGrowWidth = false
      ' TextHorizontalAdjust: maldekstra 0, centra 1, dekstra 2, kompleta 3
      ' ParaAdjust: maldekstra 0, dekstra 1, kompleta 2, centra 3, kom lasta 4
      if io(0)+io(1)+io(2)+io(3)+io(4) = 1 then
        objekto.TextHorizontalAdjust = oAlghustigo
      end if
      if io(0)+io(1)+io(2)+io(3)+io(4) > 1 then
        objekto.TextHorizontalAdjust = 3
      end if
    end if
  next

end function


' get a list of all shapes whose visibility is changed during the animation
function getAnimatedShapes(slide as Object)
     shapes = Array() ' start with an empty array
 
     if not hasAnimation(slide) then
         getAnimatedShapes = shapes
         exit function
     end if
     
    mainSequence = getMainSequence(slide)    
    clickNodes = mainSequence.createEnumeration()
    while clickNodes.hasMoreElements()
        clickNode = clickNodes.nextElement()

        groupNodes = clickNode.createEnumeration()
        while groupNodes.hasMoreElements()
            groupNode = groupNodes.nextElement()

            effectNodes = groupNode.createEnumeration()
            while effectNodes.hasMoreElements()
                effectNode = effectNodes.nextElement()

                animNodes = effectNode.createEnumeration()
                while animNodes.hasMoreElements()
                    animNode = animNodes.nextElement()
                    if isVisibilityAnimation(animNode) then
                        target = animNode.target
                        if not IsEmpty(target) then 
                          ' if we haven't seen this shape yet, add it to the array
                          if not containsObject(shapes, target) then
                            newUBound = UBound(shapes) + 1
                            reDim preserve shapes(newUBound)
                            shapes(newUBound) = target
                          end if
                        end if
                     end if
                 wend
            wend
        wend
    wend
    getAnimatedShapes = shapes
end function


' create a 2-D array giving the visibility of each animated
' shape for each frame in the expanded animation
function getShapeVisibility(slide as Object, nFrames as Integer)
     shapes = getAnimatedShapes(slide)
     dim visibility(UBound(shapes), nFrames-1) as Boolean
     
     ' loop over all animated shapes
     for n = 0 to UBound(shapes)
         shape = shapes(n)
         visKnown = false
         visCurrent = false
     
         ' iterate over the animations for this slide,
         ' looking for those that change the visibility
         ' of the current shape
         mainSequence = getMainSequence(slide)
        if HasUnoInterfaces(mainSequence, ENUMACCESS) then            
            clickNodes = mainSequence.createEnumeration()
            currentFrame = 0
            while clickNodes.hasMoreElements()
                clickNode = clickNodes.nextElement()
     
                groupNodes = clickNode.createEnumeration()
                while groupNodes.hasMoreElements()
                    groupNode = groupNodes.nextElement()
     
                    effectNodes = groupNode.createEnumeration()
                    while effectNodes.hasMoreElements()
                        effectNode = effectNodes.nextElement()
     
                         animNodes = effectNode.createEnumeration()
                         while animNodes.hasMoreElements()
                             animNode = animNodes.nextElement()
                             if isVisibilityAnimation(animNode) then
                                target = animNode.target
                                ' if this is the shape we want, check the visibility
                                sameStruct = false
                                if IsUnoStruct(target) AND IsUnoStruct(shape) then
	                                sameStruct = EqualUnoObjects(shape.Shape, target.Shape) AND shape.Paragraph=target.Paragraph
                                end if
                                if EqualUnoObjects(shape, target) OR sameStruct then
                                    visCurrent = animNode.To
                                    ' if this is the first time we've seen this
                                    ' shape, set the visibility on the previous frames
                                    if visKnown = false then
                                        for i = 0 to currentFrame
                                            visibility(n, i) = not visCurrent
                                        next
                                        visKnown = true
                                    end if
                                end if
                             end if
                         wend
                     wend
                wend
                currentFrame = currentFrame + 1
                visibility(n, currentFrame) = visCurrent
            wend
        end if
    next
    getShapeVisibility = visibility
end function


' remove from the given slide all shapes that are invisible in the specified frame
sub removeInvisibleShapes(slide as Object, visibility, frame as Integer)
     shapes = getAnimatedShapes(slide)
     for n = 0 to UBound(shapes)
        if visibility(n, frame) = false then
            'special handling for com.sun.star.presentation.ParagraphTarget
            if (IsUnoStruct(shapes(n))) then
                shape=shapes(n).Shape
                para = shapes(n).Paragraph
                count = 0
                eNum =  shape.createEnumeration
                ' for each paragraph in textbox
                Do while eNum.HasMoreElements()
                    oAObj = eNum.nextElement()
                    ' remove those whose index larger than shapes(n).Paragraph
                    if count = para then
                        oAObj.String = ""
                        oAObj.NumberingIsNumber = false
                    end if
                    count = count + 1
                Loop
            else
                slide.remove(shapes(n))
            end if
        end if
    next
end sub


' checks if the given animation node changes a shape's visibility
function isVisibilityAnimation(animNode as Object) as Boolean
    isVisibilityAnimation = HasUnoInterfaces(animNode, ANIMSET) and _
                            (animNode.AttributeName = VISATTR)
end function


' check if an object (needle) is contained in an array (haystack)
function containsObject(haystack as Object, needle as Object) as Boolean
    containsObject = false
    for each item in haystack
        if EqualUnoObjects(item, needle) then
            containsObject = true
            exit function
        end if
    next item
end function


' determine whether the given drawPage has animations attached
function hasAnimation(slide as Object) as Boolean
    mainSequence = getMainSequence(slide)
    hasAnimation = HasUnoInterfaces(mainSequence, ENUMACCESS)
end function


' count the number of frames in the animation for the given slide
function countAnimationSteps(slide as Object) as Integer
    mainSequence = getMainSequence(slide)
    countAnimationSteps = countElements(mainSequence) + 1
end function


' count the number of elements in an enumerable object
function countElements(enumerable as Object) as Integer
    oEnum = enumerable.createEnumeration()
    n = 0
    while oEnum.hasMoreElements()
        n = n + 1
        oEnum.nextElement()
    wend
    countElements = n
end function


' make n-1 copies of the given slide in the given doc
' when done, the slide will appear a total of n times
' note that doc.duplicate adds copies immediately after
' the page being copied
function replicateSlide(doc, slide, n)
    for i = 1 to n-1
        doc.duplicate(slide)
    next
end function


' get the main sequence from the given draw page
function getMainSequence(oPage as Object) as Object
    on error resume next
    mainSeq = com.sun.star.presentation.EffectNodeType.MAIN_SEQUENCE

    oNodes = oPage.AnimationNode.createEnumeration()
    while oNodes.hasMoreElements()
        oNode = oNodes.nextElement()
        if getNodeType(oNode) = mainSeq then
            getMainSequence = oNode
            exit function
        end if
    wend
end function


' get the type of a node
function getNodeType(oNode as Object) as Integer
    on error resume next
    for each oData in oNode.UserData
        if oData.Name = "node-type" then
            getNodeType = oData.Value
            exit function
        end if
    next oData
end function


' get the class of an effect
function getEffectClass(oEffectNode as Object) as String
    on error resume next
    for each oData in oEffectNode.UserData
        if oData.Name = "preset-class" then
            getEffectClass = oData.Value
            exit function
        end if
    next oData
End Function


' get the id of an effect
function getEffectId(oEffectNode as Object) as String
    on error resume next
    for each oData in oEffectNode.UserData
        if oData.Name = "preset-id" then
            getEffectId = oData.Value
            exit function
        end if
    next oData
end function
