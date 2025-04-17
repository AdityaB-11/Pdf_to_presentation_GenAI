
Sub CreatePresentation()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim para As TextRange
    
    ' Create a new presentation
    Set ppt = Application.Presentations.Add

    ' Add title slide
    Set sld = ppt.Slides.Add(1, ppLayoutTitle)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Huffman Coding: Data Compression Technique"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Add creator name if provided
    If sld.Shapes.HasTitle Then
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 400, 600, 50)
        Set tf = shp.TextFrame
        tf.TextRange.Text = "Created by: "
        tf.TextRange.Font.Size = 14
        tf.TextRange.Font.Color.RGB = RGB(128, 128, 128)  ' Gray color
        tf.HorizontalAlignment = ppAlignCenter
    End If

    ' Add index slide
    Set sld = ppt.Slides.Add(2, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Index"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.MarginLeft = 20
    tf.MarginRight = 20

    ' Add index content

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "1."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Introduction to Huffman Coding"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "2."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "How Huffman Coding Works: An Example"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "3."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Building the Huffman Tree"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "4."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman Coding Algorithm & Code Implementation (Snippet)"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "5."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Prefix Codes and Ambiguity Prevention"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "6."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Encoding and Decoding Example"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "7."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion: Advantages & Disadvantages of Huffman Coding"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0

    ' Add blank line before conclusion
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = ""
    para.ParagraphFormat.SpaceAfter = 6
    
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceBefore = 6

    ' Add slide 3
    Set sld = ppt.Slides.Add(3, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Introduction to Huffman Coding"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Developed by David Huffman."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "A lossless data compression technique."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Reduces data size without losing information."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Most effective for data with frequently occurring characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Uses variable-length codes based on character frequency."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Creates a binary tree to represent character codes."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Achieves compression by assigning shorter codes to more frequent characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 4
    Set sld = ppt.Slides.Add(4, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "How Huffman Coding Works: An Example"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Consider a string needing transmission (e.g., ""ABCCCBBAA"")."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Each character typically uses 8 bits (1 byte)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Calculating total bits for the initial string (15 characters * 8 bits/character = 120 bits)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman coding aims to reduce this bit count."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "It achieves this by assigning shorter codes to frequent characters and longer codes to less frequent characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This leads to a smaller overall size for the encoded data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 5
    Set sld = ppt.Slides.Add(5, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Building the Huffman Tree"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Calculate the frequency of each character in the input string."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Sort characters in ascending order of frequency (using a priority queue)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Create a leaf node for each character, with its frequency as a value."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Repeatedly combine the two nodes with the lowest frequencies:"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Create a new parent node."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Assign the sum of the children's frequencies to the parent."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Assign '0' to the left branch and '1' to the right branch."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Continue until only one node (the root) remains."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 6
    Set sld = ppt.Slides.Add(6, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Huffman Coding Algorithm & Code Implementation (Snippet)"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Priority Queue:  Use a Min-Heap data structure to efficiently manage character frequencies."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Tree Construction:  Iteratively build the Huffman tree by merging nodes with the smallest frequencies."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Code Generation: Traverse the tree to assign binary codes to each character (0 for left, 1 for right)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Code Example (C/C++): Show snippets of functions like `buildHuffmanTree`, `printHCodes`, etc."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "(Include relevant data ..."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Illustrate the key steps of the algorithm within the code."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 7
    Set sld = ppt.Slides.Add(7, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Prefix Codes and Ambiguity Prevention"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman codes are prefix codes."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "No code is a prefix of another code."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This prevents ambiguity during decoding."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The tree structure ensures this prefix property."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Decoding is unambiguous because you can uniquely traverse the tree based on the binary code."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 8
    Set sld = ppt.Slides.Add(8, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Encoding and Decoding Example"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Show a table with characters, their frequencies, and assigned Huffman codes."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Demonstrate the encoding process: replace each character in the input string with its Huffman code."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Calculate the size of the encoded string."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Show how to decode the encoded string using the Huffman tree."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Illustrate the decoding process step-by-step, using an example code."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 9
    Set sld = ppt.Slides.Add(9, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Conclusion: Advantages & Disadvantages of Huffman Coding"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Advantages:  Lossless compression, efficient for data with repetitive characters, relatively simple algorithm."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Disadvantages:  Overhead of transmitting the Huffman tree, not optimal for all types of data, can be less efficient t..."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Applications: Data compression in various fields, text compression, image compression (in conjunction with other tech..."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Future improvements and related algorithms could be mentioned briefly."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

End Sub
