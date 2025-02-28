
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
    sld.Shapes.Title.TextFrame.TextRange.Text = "Huffman Coding: Data Compression Through Efficient Encoding"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Add creator name if provided
    If sld.Shapes.HasTitle Then
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 400, 600, 50)
        Set tf = shp.TextFrame
        tf.TextRange.Text = "Created by: guppi"
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
    para.Text = "Huffman Coding Algorithm: A Step-by-Step Guide"
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
    para.Text = "Encoding and Decoding Example & Code Implementation Overview"
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
    para.Text = "Conclusion: Efficiency and Applications of Huffman Coding"
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
    para.Text = "Huffman coding is a lossless data compression algorithm that reduces file size by assigning shorter codes to more frequent characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "It was developed by David A. Huffman and published in 1952."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The algorithm is based on the frequency of characters in the input data.  More frequent characters receive shorter codes, and less frequent characters receive longer codes."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This variable-length coding scheme leads to significant compression, especially for text with repetitive characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman coding is widely used in various applications, including data archiving, file compression (e.g., .zip, .rar), and fax transmission."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "It forms the basis for many more advanced compression techniques."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Unlike fixed-length coding (e.g., ASCII), Huffman coding dynamically adjusts code lengths based on character frequency."
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
    para.Text = "Consider a string: ""ABACCABADDCAB"". Each character typically uses 8 bits (1 byte).  A 15-character string requires 120 bits (15 bytes)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman coding aims to represent this string using fewer bits."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "We first count the frequency of each character: A (5), B (2), C (3), D (2)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The algorithm then constructs a binary tree based on these frequencies, assigning shorter codes to more frequent characters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The final encoded string will be significantly shorter than the original 120 bits."
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
    para.Text = "The process begins by creating a min-heap data structure, containing nodes representing each character and its frequency."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Repeatedly, the two nodes with the lowest frequencies are extracted from the heap."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "A new node is created, its frequency being the sum of the frequencies of the extracted nodes.  The extracted nodes become the left and right children of this new node."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This new node is inserted back into the heap."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This process continues until only one node remains – the root of the Huffman tree."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Each leaf node represents a character, and the path from the root to a leaf node defines the code for that character (0 for left, 1 for right)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 6
    Set sld = ppt.Slides.Add(6, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Huffman Coding Algorithm: A Step-by-Step Guide"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 1:** Calculate the frequency of each unique character in the input string."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 2:** Create a min-heap data structure to store characters and their frequencies."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 3:**  Repeatedly extract the two nodes with the smallest frequencies from the heap."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 4:** Create a new node with these two nodes as children, the frequency of the new node being the sum of the children's frequencies. Insert the new node back into the heap."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 5:** Repeat Step 3 and 4 until only one node remains (the root of the Huffman tree)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 6:** Assign '0' to the left branch and '1' to the right branch of each node."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "**Step 7:** Traverse the tree to generate the Huffman code for each character."
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
    para.Text = "Huffman codes are prefix codes, meaning no code is a prefix of another. This prevents ambiguity during decoding."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The tree structure ensures that no code can be confused with another.  For example, if 'A' is encoded as '0' and 'B' is encoded as '01', there's no ambiguity."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This prefix property is crucial for efficient and unambiguous decoding."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The tree structure itself needs to be transmitted along with the encoded data for proper decoding. However, efficient methods exist to represent the tree concisely."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The overhead of transmitting the tree is usually small compared to the compression achieved."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 8
    Set sld = ppt.Slides.Add(8, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Encoding and Decoding Example & Code Implementation Overview"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Illustrate an encoding and decoding example using a sample string and its Huffman tree. Show the original string, frequency table, Huffman tree, generated codes, encoded string, and the decoding process."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Briefly show a high-level overview of the code (C/C++ or Python) used to implement the Huffman coding algorithm, highlighting key functions like building the tree, encoding, and decoding.  Focus on the data structures used (min-heap, tree nodes).  Avoid showing extensive code snippets."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 9
    Set sld = ppt.Slides.Add(9, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Conclusion: Efficiency and Applications of Huffman Coding"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman coding provides a highly efficient way to compress data, especially for text and other data with frequently occurring symbols."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The compression ratio depends on the distribution of character frequencies in the input data."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "It's a lossless compression method; no information is lost during compression or decompression."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Huffman coding is widely applied in various fields, including data compression utilities, image and audio compression, and network communication."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "While it's not the most advanced compression algorithm, it provides a solid foundation for understanding more sophisticated techniques.  It's often used as a building block in more complex compression methods."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

End Sub
