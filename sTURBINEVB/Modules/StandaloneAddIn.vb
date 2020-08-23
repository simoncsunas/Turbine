Module StandaloneAddIn

    'To create an instance of the SOLIDWORKS software, your project should contain lines of code similar to the following:
    Sub Main()

        Dim swApp As SldWorks.SldWorks
        swApp = New SldWorks.SldWorks()

        swApp.ExitApp()

        swApp = Nothing

        '+ Microsoft VBA-enabled Applications
        ' Attach to active Excel object
        'xl = GetObject(, "Excel.Application")

        ' Get active sheet in Excel
        'xlsh = xl.ActiveSheet

        ' Get value in Excel cell A1
        'density = xlsh.Cells(1, 1)

        ' Set density in SOLIDWORKS part
        'Part.SetUserPreferenceDoubleValue(swMaterialPropertyDensity, density)
        '- Microsoft VBA-enabled Applications

        Dim swFeature As SldWorks.Feature
        swFeature = swSketchBlockDefinition.GetFeature
        Debug.Print swFeature.Name & [" & swFeature.GetTypeName2 & "]"

        '+ Early and Late Binding
        'To implement early binding in the SOLIDWORKS software, you must reference two type libraries:
        'SldWorks version Type Library (sldworks.tlb)
        'SOLIDWORKS version Constant type library (swconst.tlb)

        Dim swApp As Object
        Dim swApp As SldWorks.SldWorks

        Dim swModel As Object
        Dim swModel As SldWorks.ModelDoc2

        Dim swEntity As Object
        Dim swEntity As SldWorks.Entity
        '- Early and Late Binding

    End Sub

    '--------------------------------------------------
    ' Preconditions: Drawing document is open, and
    ' a drawing view containing at least one component
    ' is selected.
    '
    ' Postconditions: None
    '--------------------------------------------------
    Sub DrawingViewsandModelEntities()

        Dim swApp As SldWorks.SldWorks
        Dim swModel As SldWorks.ModelDoc2
        Dim swSelMgr As SldWorks.SelectionMgr
        Dim swDrawing As SldWorks.DrawingDoc
        Dim drView As SldWorks.View
        Dim Comp As SldWorks.Component2
        Dim selData As SldWorks.SelectData
        Dim ent As SldWorks.Entity
        Dim EntComp As SldWorks.Component2

        Dim itr As Long
        Dim CompCount As Long

        Dim vComps As Object
        Dim vEdges As Object
        Dim vVerts As Object
        Dim vFaces As Object

        Dim i As Long

        Dim boolstatus As Boolean

        swApp = Application.SldWorks
        swModel = swApp.ActiveDoc
        swDrawing = swModel
        swSelMgr = swModel.SelectionManager
        drView = swDrawing.ActiveDrawingView

        Debug.Assert(Not drView Is Nothing)

        Debug.Print "Name of drawing view: "; drView.Name
        Debug.Print()

        CompCount = drView.GetVisibleComponentCount

        Debug.Assert(CompCount <> 0)
        Debug.Print "Number of visible components = "; CompCount

        vComps = drView.GetVisibleComponents

        Debug.Assert(Not IsEmpty(vComps))

        For i = LBound(vComps) To UBound(vComps)

            swModel.ClearSelection2(True)

            Debug.Print("")
            Debug.Print("Component " & i & " name is " & vComps(i).Name2)

            Comp = vComps(i)

            'Get all edges of this component that are visible in this drawing view
            vEdges = drView.GetVisibleEntities(Comp, swViewEntityType_Edge)

            selData = swSelMgr.CreateSelectData
            selData.View = drView

            If IsEmpty(vEdges) Then

                Debug.Print("   No edges")

            Else

                Debug.Print("   This component has " & UBound(vEdges) + 1 & " visible edges in this view.")

                For itr = 0 To UBound(vEdges)

                    ent = vEdges(itr)
                    boolstatus = ent.Select4(False, selData)

                Next itr

            End If


            'Get all vertices of this component that are visible in this drawing view
            vVerts = drView.GetVisibleEntities(Comp, swViewEntityType_Vertex)

            If IsEmpty(vVerts) Then

                Debug.Print("   No vertices")

            Else

                Debug.Print("   This component has " & UBound(vVerts) + 1 & " visible vertices in this view")

                For itr = 0 To UBound(vVerts)

                    ent = vVerts(itr)
                    boolstatus = ent.Select4(False, selData)

                Next itr

            End If


            swModel.ClearSelection2(True)

            'Get all faces of this component that are visible in this drawing view
            vFaces = drView.GetVisibleEntities(Comp, swViewEntityType_Face)

            If IsEmpty(vFaces) Then

                Debug.Print("   No faces")

            Else

                Debug.Print("   This component has " & UBound(vFaces) + 1 & " visible faces in this view.")

                For itr = 0 To UBound(vFaces)

                    ent = vFaces(itr)
                    boolstatus = ent.Select4(False, selData)

                Next itr

            End If



        Next i

        'Get all the entities (edges, faces, and vertices) that are visible in the drawing view
        Debug.Print()

        swModel.ClearSelection2(True)

        Comp = Nothing

        EntComp = Nothing

        'Get all edges of all components that are visible in this drawing view
        vEdges = drView.GetVisibleEntities(Comp, swViewEntityType_Edge)

        Debug.Print("There are a total of " & UBound(vEdges) + 1 & " visible edges in this view.")

        For itr = 0 To UBound(vEdges)

            ent = vEdges(itr)

            'boolstatus = ent.Select4(False, selData)
            EntComp = ent.GetComponent

            EntComp = Nothing

        Next itr


        'Get all vertices of all components that are visible in this drawing view
        vVerts = drView.GetVisibleEntities(Comp, swViewEntityType_Vertex)

        Debug.Print("There are a total of " & UBound(vVerts) + 1 & " visible vertices in this view.")

        For itr = 0 To UBound(vVerts)

            ent = vVerts(itr)

            'boolstatus = ent.Select4(False, selData)
            EntComp = ent.GetComponent

            EntComp = Nothing

        Next itr

        swModel.ClearSelection2(True)

        'Get all faces of all components that are visible in this drawing view
        vFaces = drView.GetVisibleEntities(Comp, swViewEntityType_Face)

        Debug.Print("There are a total of " & UBound(vFaces) + 1 & " visible faces in this view.")

        For itr = 0 To UBound(vFaces)

            ent = vFaces(itr)
            'boolstatus = ent.Select4(False, selData)

        Next itr
    End Sub

End Module
