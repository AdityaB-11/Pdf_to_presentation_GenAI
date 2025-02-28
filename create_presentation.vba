
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
    sld.Shapes.Title.TextFrame.TextRange.Text = "Polymer Chemistry: Types, Properties, and Biomedical Applications"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Add creator name if provided
    If sld.Shapes.HasTitle Then
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 400, 600, 50)
        Set tf = shp.TextFrame
        tf.TextRange.Text = "Created by: shreyash katole"
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
    para.Text = "Condensation Polymerization: Step-Growth Reaction"
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
    para.Text = "Polymer Applications: Examples in Daily Life"
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
    para.Text = "Conducting Polymers: Electrical Conductivity in Polymers"
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
    para.Text = "Types of Conducting Polymers and Doping"
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
    para.Text = "Polymers in Medicine and Surgery: Biomaterials"
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
    para.Text = "Conclusion: The Versatile World of Polymers"
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
    sld.Shapes.Title.TextFrame.TextRange.Text = "Condensation Polymerization: Step-Growth Reaction"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Condensation polymerization requires monomers with at least two functional groups capable of reacting with each other."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "These functional groups react, forming a bond between monomers and releasing a small molecule as a byproduct (e.g., water, ammonia, hydrogen chloride)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Mineral acids or bases commonly serve as catalysts in condensation polymerization."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Unlike addition polymerization, this process is endothermic, absorbing heat."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "It is generally a slower process than addition polymerization."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Condensation polymers typically have higher molecular weights than addition polymers."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This method often results in three-dimensional, cross-linked, thermosetting polymers."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 4
    Set sld = ppt.Slides.Add(4, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Polymer Applications: Examples in Daily Life"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Polyethylene is widely used in disposable syringes due to its flexibility and low cost."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Polypropylene's strength and biocompatibility make it suitable for heart walls and blood filters."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Polyvinyl chloride (PVC) also finds application in disposable syringes."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Acrylic hydrogels are employed in grafting procedures due to their biocompatibility and water-absorbing properties."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Polymethyl methacrylate (PMMA) is a common material for contact lenses owing to its transparency and biocompatibility."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Poly(alkyl sulfone) membranes are used in oxygenators for their gas permeability and biocompatibility."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 5
    Set sld = ppt.Slides.Add(5, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Conducting Polymers: Electrical Conductivity in Polymers"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Most polymers are electrical insulators because their electrons are tightly bound in sigma bonds, limiting mobility."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conductivity is achieved in polymers through the presence of conjugated (alternating) double bonds, creating delocalized pi electrons."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "These delocalized pi electrons can move relatively freely along the polymer chain under an electric field, enabling electrical conduction."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "This creates a system of valence and conduction bands, similar to those in metals.  Examples include polyacetylene, polyanilines, and polythiophene."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The conductivity of these polymers can be further enhanced through doping."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 6
    Set sld = ppt.Slides.Add(6, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Types of Conducting Polymers and Doping"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Intrinsically conducting polymers (ICPs) possess delocalized electrons within their backbone structure (e.g., polyacetylene, polypyrrole, polythiophene)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Extrinsically conducting polymers achieve conductivity through added components: filled polymers (conducting elements like carbon black dispersed in a polymer matrix) and blended polymers (blending conventional and conducting polymers)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Doping increases conductivity by introducing charge carriers: p-doping (oxidation using Lewis acids, creating positive charges) and n-doping (reduction using Lewis bases, creating negative charges)."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "P-doping examples include treatment with iodine or iron(III) chloride."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "N-doping examples include treatment with lithium or sodium."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Coordination conducting polymers are formed by combining a metal atom with a polydentate ligand."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 7
    Set sld = ppt.Slides.Add(7, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Polymers in Medicine and Surgery: Biomaterials"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Biomaterials are materials used in medical applications without causing adverse effects.  Polymer biomaterials are increasingly important."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Key characteristics of biomedical polymers include biocompatibility (lack of harmful reactions with body tissues), purity and reproducibility, and sterilizability without property alteration."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "They must possess optimal physical and chemical properties for their intended use."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Biomedical polymers must not destroy blood cellular elements or enzymes, nor cause toxic or allergic reactions."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "They should also maintain electrolyte balance in the body."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    ' Add slide 8
    Set sld = ppt.Slides.Add(8, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Conclusion: The Versatile World of Polymers"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Polymers exhibit a wide range of properties and applications, from everyday plastics to advanced biomaterials and conducting polymers."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Understanding the different types of polymerization and the factors influencing polymer properties is crucial for designing materials with specific functionalities."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Continued research and development in polymer chemistry will lead to further innovations in various fields, including medicine, electronics, and materials science."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "The versatility of polymers makes them essential materials for modern technology and improving the quality of life."
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14

End Sub
