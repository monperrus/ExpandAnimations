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
Private LAST_UNSUPPORTED_SLIDES as Integer


' Expands the current document and writes filename-expanded.odp plus filename.pdf.
sub Main 
  Dim doc As Object
  doc = thisComponent
  
  newUrlExpanded = getExpandedOdpUrl(doc)

  ' the expansion
  newUrlPdf = expandAnimations(doc)

  message = "Expansion done!" + Chr(10) + "ODP: " + newUrlExpanded + Chr(10) + "PDF: " + newUrlPdf
  if LAST_UNSUPPORTED_SLIDES > 0 then
    message = message + Chr(10) + Chr(10) + "Warning: " + CStr(LAST_UNSUPPORTED_SLIDES) + " slide(s) contain unsupported animation effects that were not expanded."
  end if
  msgbox message
end sub

' expands the current document without showing a dialog
sub Headless(Optional sourceUrl as String)
  hasSourceUrl = false
  if not IsMissing(sourceUrl) then
    if sourceUrl <> "" then
      hasSourceUrl = true
    end if
  end if

  if not hasSourceUrl then
    doc = getActiveDocument()
    expandAnimations(doc)
  else
    doc = openDocument(sourceUrl)
    expandAnimations(doc, getDocumentUrlFromSource(sourceUrl))
  end if
end sub

' expands the current document without showing a dialog, then exits LibreOffice
sub CommandLine(Optional sourceUrl as String)
  hasSourceUrl = false
  if not IsMissing(sourceUrl) then
    if sourceUrl <> "" then
      hasSourceUrl = true
    end if
  end if

  if not hasSourceUrl then
    sourceUrl = Environ("EXPANDANIMATIONS_INPUT")
  end if

  if sourceUrl = "" then
    Headless
  else
    Headless(sourceUrl)
  end if
  StarDesktop.terminate
end sub


function getActiveDocument()
  On Error Resume Next

  for attempt = 1 to 50
    doc = StarDesktop.CurrentComponent
    docUrl = getDocumentUrl(doc)
    if docUrl <> "" then
      getActiveDocument = doc
      exit function
    end if
    Wait 200
  next

  getActiveDocument = ThisComponent
end function


function getDocumentUrl(doc as Object) as String
  On Error Resume Next
  getDocumentUrl = ""
  getDocumentUrl = doc.getURL()
end function


function openDocument(sourceUrl as String)
  docUrl = getDocumentUrlFromSource(sourceUrl)

  Dim loadProps()
  openDocument = StarDesktop.loadComponentFromURL(docUrl, "_blank", 0, loadProps)
end function


function getDocumentUrlFromSource(sourceUrl as String) as String
  getDocumentUrlFromSource = sourceUrl
  if Left(getDocumentUrlFromSource, 5) <> "file:" then
    getDocumentUrlFromSource = ConvertToURL(sourceUrl)
  end if
end function

' tests the module
' can be called on the command line with
' $ EXPANDANIMATIONS_INPUT=${pathToTestFile} libreoffice --headless "macro:///ExpandAnimations.ExpandAnimations.test"
sub test
  CommandLine
end sub

' expands the animations and exports to ODP and PDF
function expandAnimations(doc as Object, Optional sourceUrl as String)
  if IsMissing(sourceUrl) then
    sDocUrl = doc.getURL()
  else
    sDocUrl = sourceUrl
  end if
  sDocPath = getDirectoryNameFromUrl(sDocUrl)
  sDocFileNameWithoutExtension = getFileNameWithoutExtensionFromUrl(sDocUrl)
  newUrlExpanded = getExpandedOdpUrlFromUrl(sDocUrl)
  newUrlPdf = sDocPath + "/" + sDocFileNameWithoutExtension + ".pdf"

  if IsMissing(sourceUrl) then
    docExpanded = renameAsExpanded(doc, newUrlExpanded)
  else
    docExpanded = renameAsExpanded(doc, newUrlExpanded, sDocUrl)
  end if

  expandDocument(docExpanded)
  exportToPDF(docExpanded, newUrlPdf)

  expandAnimations = newUrlPdf
  
  docExpanded.close(false)
end function


' returns the URL used for the expanded ODP
function getExpandedOdpUrl(doc as Object)
  getExpandedOdpUrl = getExpandedOdpUrlFromUrl(doc.getURL())
end function


function getExpandedOdpUrlFromUrl(sDocUrl as String)
  sDocPath = getDirectoryNameFromUrl(sDocUrl)
  sDocFileNameWithoutExtension = getFileNameWithoutExtensionFromUrl(sDocUrl)

  getExpandedOdpUrlFromUrl = sDocPath + "/" + sDocFileNameWithoutExtension + "-expanded.odp"
end function


function getDirectoryNameFromUrl(docUrl as String) as String
  slashPos = lastIndexOf(docUrl, "/")
  if slashPos > 0 then
    getDirectoryNameFromUrl = Left(docUrl, slashPos-1)
  else
    getDirectoryNameFromUrl = docUrl
  end if
end function


function getFileNameWithoutExtensionFromUrl(docUrl as String) as String
  slashPos = lastIndexOf(docUrl, "/")
  if slashPos > 0 then
    fileName = Mid(docUrl, slashPos+1)
  else
    fileName = docUrl
  end if

  dotPos = lastIndexOf(fileName, ".")
  if dotPos > 1 then
    getFileNameWithoutExtensionFromUrl = Left(fileName, dotPos-1)
  else
    getFileNameWithoutExtensionFromUrl = fileName
  end if
end function


function lastIndexOf(value as String, token as String) as Integer
  lastIndexOf = 0
  tokenLen = Len(token)
  if tokenLen = 0 then
    exit function
  end if

  for pos = Len(value) - tokenLen + 1 to 1 step -1
    if Mid(value, pos, tokenLen) = token then
      lastIndexOf = pos
      exit function
    end if
  next
end function


' saves the (current) document with a new file name
' e.g. test.odp -> test-expanded.odp
function renameAsExpanded(doc as Object, newUrlExpanded as String, Optional sourceUrl as String)
  Dim Dummy()

  if IsMissing(sourceUrl) then
    doc.storeToUrl(newUrlExpanded, Array(makePropertyValue("Overwrite", True)))
  else
    fileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
    if fileAccess.exists(newUrlExpanded) then
      fileAccess.kill(newUrlExpanded)
    end if
    fileAccess.copy(sourceUrl, newUrlExpanded)
  end if
  
  expandedDoc = StarDesktop.loadComponentFromURL(newUrlExpanded, "_default", 0, Dummy)  
  
  renameAsExpanded = expandedDoc
end function


' exports to PDF
sub exportToPDF(doc as Object, newUrlPdf as String)
  Dim filterData(3)

  filterData(0) = makePropertyValue("ExportBookmarksToPDFDestination", True)
  filterData(1) = makePropertyValue("PDFViewSelection", 1)
  filterData(2) = makePropertyValue("ConvertOOoTargetToPDFTarget", True)
  filterData(3) = makePropertyValue("EmbedStandardFonts", True)

  doc.storeToUrl(newUrlPdf, Array(makePropertyValue("FilterName", "impress_pdf_Export"), makePropertyValue("FilterData", filterData)))
end sub



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
    LAST_UNSUPPORTED_SLIDES = 0
    
    numSlides = doc.getDrawPages().getCount()
    sourcePdfPages = getSourcePdfPageMap(doc)

    ' Work backwards so duplicated slides do not disturb the remaining indexes.
    for i = numSlides-1 to 0 step -1
        slide = doc.drawPages(i)
        fixatePageFieldsInShapes(slide, i+1, numSlides)
        if slide.IsPageNumberVisible then
            fixateMasterPageNumber(doc, slide, i+1, numSlides)
        end if
        ' Auto-generated slide names are not stable PDF targets.
        if Left(slide.Name, 4) = "page" then
        	slide.Name = "Slide: " & CStr(i+1)
        end if
        if hasAnimation(slide) then
            n = countAnimationSteps(slide)
            if n > 1 and hasUnsupportedAnimation(slide) then
                LAST_UNSUPPORTED_SLIDES = LAST_UNSUPPORTED_SLIDES + 1
            end if
            if n > 1 and hasNoSupportedAnimationTargets(slide) then
                ' keep slide as-is; unsupported effects are reported above
            elseif n > 1 then
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
        Wait 1
    next

    removeHiddenSlides(doc)
    renameSlidesForExportedOrder(doc)
    fixInternalLinks(doc, sourcePdfPages)
    clearAllAnimations(doc)
    
  storeExpandedDocument(doc)
end function


sub storeExpandedDocument(doc as Object)
    doc.storeToUrl(doc.getURL(), Array(makePropertyValue("Overwrite", True)))
end sub


' Name visible slides after their exported PDF page number so internal links
' survive both the macro PDF export and later manual exports from the ODP.
sub renameSlidesForExportedOrder(doc as Object)
    pages = doc.getDrawPages()

    for pageNr = 0 to pages.getCount()-1
        pages.getByIndex(pageNr).Name = "ExpandAnimationsTmp" & CStr(pageNr+1)
    next

    pdfPage = 1
    hiddenPage = 1
    for pageNr = 0 to pages.getCount()-1
        slide = pages.getByIndex(pageNr)
        if slide.Visible then
            slide.Name = "Slide: " & CStr(pdfPage)
            pdfPage = pdfPage + 1
        else
            slide.Name = "Hidden Slide: " & CStr(hiddenPage)
            hiddenPage = hiddenPage + 1
        end if
    next
end sub


sub removeHiddenSlides(doc as Object)
    pages = doc.getDrawPages()
    for pageNr = pages.getCount()-1 to 0 step -1
        slide = pages.getByIndex(pageNr)
        if not slide.Visible then
            pages.remove(slide)
        end if
    next
end sub


' Map original slide numbers to their first visible page in the exported PDF.
function getSourcePdfPageMap(doc as Object)
    numSlides = doc.getDrawPages().getCount()
    Dim sourcePdfPages() As Integer
    ReDim sourcePdfPages(numSlides-1)
    pdfPage = 1

    for slideNr = 0 to numSlides-1
        slide = doc.drawPages(slideNr)
        if slide.Visible then
            sourcePdfPages(slideNr) = pdfPage
            pdfPage = pdfPage + getExpandedFrameCount(slide)
        else
            sourcePdfPages(slideNr) = 0
        end if
    next

    getSourcePdfPageMap = sourcePdfPages
end function


function getExpandedFrameCount(slide as Object) as Integer
    getExpandedFrameCount = 1
    if hasAnimation(slide) then
        n = countAnimationSteps(slide)
        if n > 1 and not hasNoSupportedAnimationTargets(slide) then
            getExpandedFrameCount = n
        end if
    end if
end function


' Rewrite internal links to target the visible PDF page produced by expansion.
sub fixInternalLinks(doc as Object, sourcePdfPages)
    pages = doc.getDrawPages()
    for pageNr = 0 to pages.getCount()-1
        fixInternalLinksInShapes(doc, pages.getByIndex(pageNr), sourcePdfPages)
    next
end sub


sub fixInternalLinksInShapes(doc as Object, shapes as Object, sourcePdfPages)
    for shapeNr = 0 to shapes.getCount()-1
        shape = shapes.getByIndex(shapeNr)
        if HasUnoInterfaces(shape, "com.sun.star.container.XIndexAccess") then
            fixInternalLinksInShapes(doc, shape, sourcePdfPages)
        elseif HasUnoInterfaces(shape, "com.sun.star.text.XText") then
            fixInternalLinksInText(doc, shape.Text, sourcePdfPages)
        end if
    next
end sub


sub fixInternalLinksInText(doc as Object, text as Object, sourcePdfPages)
    paragraphs = text.createEnumeration()
    while paragraphs.hasMoreElements()
        paragraph = paragraphs.nextElement()
        portions = paragraph.createEnumeration()
        while portions.hasMoreElements()
            fixInternalLinkInPortion(doc, portions.nextElement(), sourcePdfPages)
        wend
    wend
end sub


sub fixInternalLinkInPortion(doc as Object, portion as Object, sourcePdfPages)
    On Error Resume Next

    if portion.TextPortionType = "TextField" then
        field = portion.TextField
        if field.supportsService("com.sun.star.text.TextField.URL") then
            fixedUrl = getFixedInternalLinkUrl(field.URL, sourcePdfPages)
            if fixedUrl <> field.URL then
                replacement = doc.createInstance("com.sun.star.text.TextField.URL")
                replacement.URL = fixedUrl
                replacement.Representation = field.Representation
                replacement.TargetFrame = field.TargetFrame
                field.Anchor.Text.insertTextContent(field.Anchor, replacement, True)
            end if
        end if
    end if

    linkUrl = ""
    linkUrl = portion.HyperLinkURL
    if linkUrl <> "" then
        fixedUrl = getFixedInternalLinkUrl(linkUrl, sourcePdfPages)
        if fixedUrl <> linkUrl then
            portion.HyperLinkURL = fixedUrl
        end if
    end if
end sub


function getFixedInternalLinkUrl(url as String, sourcePdfPages)
    getFixedInternalLinkUrl = url

    sourceSlideNr = getSlideNumberFromInternalLink(url)
    if sourceSlideNr < 1 or sourceSlideNr > UBound(sourcePdfPages)+1 then
        exit function
    end if

    pdfPage = sourcePdfPages(sourceSlideNr-1)
    if pdfPage > 0 then
        getFixedInternalLinkUrl = "#Slide: " & CStr(pdfPage)
    end if
end function


function getSlideNumberFromInternalLink(url as String) as Integer
    getSlideNumberFromInternalLink = 0
    if Len(url) < 2 or Left(url, 1) <> "#" then
        exit function
    end if

    target = Mid(url, 2)
    endPos = Len(target)
    while endPos > 0 and not isAsciiDigit(Mid(target, endPos, 1))
        endPos = endPos - 1
    wend
    if endPos = 0 then
        exit function
    end if

    startPos = endPos
    while startPos > 0 and isAsciiDigit(Mid(target, startPos, 1))
        startPos = startPos - 1
    wend

    getSlideNumberFromInternalLink = CInt(Val(Mid(target, startPos+1, endPos-startPos)))
end function


function isAsciiDigit(ch as String) as Boolean
    if Len(ch) = 0 then
        isAsciiDigit = false
    else
        charCode = Asc(ch)
        isAsciiDigit = charCode >= 48 and charCode <= 57
    end if
end function


' Remove animation timelines from the expanded document.
sub clearAllAnimations(doc as Object)
    pages = doc.getDrawPages()
    for pageNr = 0 to pages.getCount()-1
        clearAnimations(pages.getByIndex(pageNr))
    next
end sub


sub clearAnimations(slide as Object)
    On Error Resume Next

    animationRoot = slide.AnimationNode
    animationNodes = animationRoot.createEnumeration()
    Do While animationNodes.hasMoreElements()
        animationRoot.removeChild(animationNodes.nextElement())
    Loop
end sub

function fixateMasterPageNumber(doc as Object, slide as Object, slideNr as Integer, slideCount as Integer)
    master = slide.MasterPage
    shapeCount = master.getCount()
    for shapeNr = 0 to shapeCount-1
        shape = master.getByIndex(shapeNr)
        shapeType = shape.getShapeType()
        if shapeType = "com.sun.star.presentation.SlideNumberShape" then
            copy = doc.createInstance("com.sun.star.drawing.TextShape")
            slide.IsPageNumberVisible = False
            slide.add(copy)
            copy.setString(getTextWithFixedPageFields(shape, slideNr, slideCount))
            copy.Style = shape.Style
            copy.Text.Style = shape.Text.Style
            copy.Text.CharHeight = shape.Text.CharHeight
            copy.Text.CharFontFamily = shape.Text.CharFontFamily
            copy.Text.CharFontName = shape.Text.CharFontName
            copy.Text.CharColor = shape.Text.CharColor
            copy.Position = shape.Position
            copy.Size = shape.Size
            copy.TextVerticalAdjust = shape.TextVerticalAdjust
            copy.TextHorizontalAdjust = com.sun.star.drawing.TextHorizontalAdjust.RIGHT
        end if
    next
end function

' Replace page-number and page-count fields with fixed values before duplicating slides.
function fixatePageFieldsInShapes(shapes as Object, slideNr as Integer, slideCount as Integer)
    for shapeNr = 0 to shapes.getCount()-1
        shape = shapes.getByIndex(shapeNr)
        if HasUnoInterfaces(shape, "com.sun.star.container.XIndexAccess") then
            fixatePageFieldsInShapes(shape, slideNr, slideCount)
        elseif HasUnoInterfaces(shape, "com.sun.star.text.XText") then
            fixatePageFieldsInText(shape.Text, slideNr, slideCount)
        end if
    next
end function

function fixatePageFieldsInText(text as Object, slideNr as Integer, slideCount as Integer)
    paragraphs = text.createEnumeration()
    while paragraphs.hasMoreElements()
        paragraph = paragraphs.nextElement()
        portions = paragraph.createEnumeration()
        while portions.hasMoreElements()
            portion = portions.nextElement()
            if portion.TextPortionType = "TextField" then
                field = portion.TextField
                if field.supportsService("com.sun.star.text.TextField.PageNumber") then
                    field.Anchor.String = CStr(slideNr)
                elseif field.supportsService("com.sun.star.text.TextField.PageCount") then
                    field.Anchor.String = CStr(slideCount)
                end if
            end if
        wend
    wend
end function

function getTextWithFixedPageFields(shape as Object, slideNr as Integer, slideCount as Integer)
    fixedText = ""
    firstParagraph = true
    paragraphs = shape.Text.createEnumeration()
    while paragraphs.hasMoreElements()
        if not firstParagraph then
            fixedText = fixedText + Chr(13)
        end if
        firstParagraph = false

        paragraph = paragraphs.nextElement()
        portions = paragraph.createEnumeration()
        while portions.hasMoreElements()
            portion = portions.nextElement()
            portionText = portion.String
            if portion.TextPortionType = "TextField" then
                field = portion.TextField
                if field.supportsService("com.sun.star.text.TextField.PageNumber") then
                    portionText = CStr(slideNr)
                elseif field.supportsService("com.sun.star.text.TextField.PageCount") then
                    portionText = CStr(slideCount)
                end if
            end if
            fixedText = fixedText + portionText
        wend
    wend

    getTextWithFixedPageFields = fixedText
end function

function hasNoSupportedAnimationTargets(slide as Object)
    shapes = getAnimatedShapes(slide)
    if UBound(shapes) = -1 then
      hasNoSupportedAnimationTargets = true
    else
      hasNoSupportedAnimationTargets = false
    end if
end function

' Get the shapes whose visibility changes during the animation.
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
                if HasUnoInterfaces(effectNode, ENUMACCESS) then
                  animNodes = effectNode.createEnumeration()
                  while animNodes.hasMoreElements()
                      animNode = animNodes.nextElement()
                      if isVisibilityAnimation(animNode) then
                          target = animNode.target
                          if not IsEmpty(target) then
                              if not containsObject(shapes, target) then
                                  newUBound = UBound(shapes) + 1
                                  reDim preserve shapes(newUBound)
                                  shapes(newUBound) = target
                              end if
                          end if
                      end if
                  wend
              end if
            wend
        wend
    wend
    getAnimatedShapes = shapes
end function


function hasUnsupportedAnimation(slide as Object) as Boolean
     hasUnsupportedAnimation = false

     if not hasAnimation(slide) then
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
                if HasUnoInterfaces(effectNode, ENUMACCESS) then
                  animNodes = effectNode.createEnumeration()
                  while animNodes.hasMoreElements()
                      animNode = animNodes.nextElement()
                      if not isVisibilityAnimation(animNode) then
                          hasUnsupportedAnimation = true
                          exit function
                      end if
                  wend
              end if
            wend
        wend
    wend
end function


' Build the visibility state of each animated shape for each exported frame.
function getShapeVisibility(slide as Object, nFrames as Integer)
     shapes = getAnimatedShapes(slide)
     dim visibility(UBound(shapes), nFrames-1) as Boolean
     
     for n = 0 to UBound(shapes)
         shape = shapes(n)
         visKnown = false
         visCurrent = false
     
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
                         if HasUnoInterfaces(effectNode, ENUMACCESS) then
                           animNodes = effectNode.createEnumeration()
                           while animNodes.hasMoreElements()
                               animNode = animNodes.nextElement()
                               if isVisibilityAnimation(animNode) then
                                  target = animNode.target
                                  sameStruct = false
                                  if IsUnoStruct(target) AND IsUnoStruct(shape) then
                                    sameStruct = EqualUnoObjects(shape.Shape, target.Shape) AND shape.Paragraph=target.Paragraph
                                  end if
                                  if EqualUnoObjects(shape, target) OR sameStruct then
                                      visCurrent = animNode.To
                                      if visKnown = false then
                                          for i = 0 to currentFrame
                                              visibility(n, i) = not visCurrent
                                          next
                                          visKnown = true
                                      end if
                                  end if
                               end if
                           wend
                       end if
                     wend
                wend
                currentFrame = currentFrame + 1
                visibility(n, currentFrame) = visCurrent
            wend
        end if
    next
    getShapeVisibility = visibility
end function


' Remove content that should not be visible in the specified exported frame.
sub removeInvisibleShapes(slide as Object, visibility, frame as Integer)
     shapes = getAnimatedShapes(slide)
     for n = 0 to UBound(shapes)
        if visibility(n, frame) = false then
            if (IsUnoStruct(shapes(n))) then
                shape=shapes(n).Shape
                para = shapes(n).Paragraph
                count = 0
                eNum =  shape.createEnumeration
                Do while eNum.HasMoreElements()
                    oAObj = eNum.nextElement()
                    ' Keep a blank line so vertical alignment and autofit stay stable.
                    if count >= para then
                        oAObj.String = " "
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


' Check whether an animation node changes a shape's visibility.
function isVisibilityAnimation(animNode as Object) as Boolean
    On Error Resume Next
    isVisibilityAnimation = False
    isVisibilityAnimation = HasUnoInterfaces(animNode, ANIMSET) and _
                            (animNode.AttributeName = VISATTR)
end function


function containsObject(haystack as Object, needle as Object) as Boolean
    containsObject = false
    for each item in haystack
        if EqualUnoObjects(item, needle) then
            containsObject = true
            exit function
        end if
    next item
end function


' Determine whether the given drawPage has animations attached.
function hasAnimation(slide as Object) as Boolean
    mainSequence = getMainSequence(slide)
    hasAnimation = HasUnoInterfaces(mainSequence, ENUMACCESS)
end function


' Count the number of frames in the animation for the given slide.
function countAnimationSteps(slide as Object) as Integer
    mainSequence = getMainSequence(slide)
    countAnimationSteps = countElements(mainSequence) + 1
end function


function countElements(enumerable as Object) as Integer
    oEnum = enumerable.createEnumeration()
    n = 0
    while oEnum.hasMoreElements()
        n = n + 1
        oEnum.nextElement()
    wend
    countElements = n
end function


' doc.duplicate inserts each copy immediately after the source slide.
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
