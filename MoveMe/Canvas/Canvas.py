# -*- coding: utf-8 -*-
import wx
import MoveMe
import json
import logging
from random import randint
from datetime import datetime
from MoveMe.Canvas.Factories.DefaultNodesFactory import DefaultNodesFactory
from MoveMe.Canvas.Factories.DefaultConnectionsFactory import DefaultConnectionsFactory
import wx.lib.sheet as sheet


#Define Text Drop Target class
class TextDropTarget(wx.TextDropTarget):
    def __init__(self, canvas):
        wx.TextDropTarget.__init__(self)
        self._canvas = canvas

    def OnDropText(self, x, y, data):
        self._canvas.ButtonAddNode(None, xy=[x,y])

class Matrix(sheet.CSheet):
    def __init__(self, parent):
        sheet.CSheet.__init__(self, parent)
        self.SetLabelBackgroundColour('#DBD4D4')

        self.SetDefaultColSize(20,20)
        self.SetNumberRows(0)
        self.SetNumberCols(0)

class Canvas(wx.PyScrolledWindow):
    """
    Canvas stores and renders all nodes and node connections.
    It also handles all user interaction.
    """

    def __init__(self, parent, tree, nodesFactory=DefaultNodesFactory(), connectionsFactory=DefaultConnectionsFactory(), **kwargs):
        super(Canvas, self).__init__(parent=parent)
        self.applicationId = "MoveMe"
        self.tree = tree

        self.delete_button = wx.Button(self, wx.ID_ANY, u"Delete", (20, 420), (80, 23), 0 )
        self.add_button = wx.Button(self, wx.ID_ANY, u"Add node", (20, 460), (80, 23), 0 )
        self.save_button = wx.Button(self, wx.ID_ANY, u"Save snapshot", (120, 420), (80, 23), 0 )
        self.load_button = wx.Button(self, wx.ID_ANY, u"Load snapshot", (120, 460), (80, 23), 0 )
        self.clear_button = wx.Button(self, wx.ID_ANY, u"Clear", (220, 460), (80, 23), 0 )
        self.new_nod = wx.TextCtrl(self, -1, u"D&D new node", (220, 420), (80, 23))
        self.new_nod.SetEditable(0)

        wx.StaticText(self, -1, 'Gpaph snapshots:', pos = (320, 420))
        self.saved_canvas_box = wx.CheckListBox(self, -1, size = (280, 100), pos=(320,440))
        wx.StaticText(self, -1, 'History:', pos = (610, 420))
        self.history_box = wx.CheckListBox(self, -1, size = (365, 100), pos=(610,440))

        self.nb = wx.Notebook(self, -1, pos=(610, -20), size=(370, 420), style=wx.NB_HITTEST_NOWHERE)
        self.sheet1 = Matrix(self.nb)
        self.nb.AddPage(self.sheet1, "")

        self.scrollStep = kwargs.get("scrollStep", 10)
        self.canvasDimensions = kwargs.get("canvasDimensions", [600, 400])
        self.SetScrollbars(self.scrollStep,
                           self.scrollStep,
                           self.canvasDimensions[0]/self.scrollStep,
                           self.canvasDimensions[1]/self.scrollStep)

        #This list stores all objects on canvas
        self._canvasObjects = []
        self._nodesFactory = nodesFactory
        self._connectionsFactory = connectionsFactory

        #References to objects required for implementing moving, highlighting, etc
        self._objectUnderCursor = None
        self._draggingObject = None
        self._lastDraggingPosition = None
        self._lastLeftDownPos = None
        self._selectedObject = None
        self._connectionStartObject = None

        self.SavedCanvas = []
        self.AllSavedCanvas = []
        self.History = []
        self.AllSavedHistory = []
        self.ItemsToDelete = []
        self.numberForNod = 1
        self.numberNods = []

        self.__allConectionsNumber = 0
        #Rendering initialization
        self._dcBuffer = wx.EmptyBitmap(*self.canvasDimensions)
        self.Render()
        self.Bind(wx.EVT_PAINT,
                  lambda evt: wx.BufferedPaintDC(self, self._dcBuffer, wx.BUFFER_VIRTUAL_AREA)
                  )

        #User interaction handling
        self.Bind(wx.EVT_MOTION, self.OnMouseMotion)
        self.Bind(wx.EVT_LEFT_DOWN, self.OnMouseLeftDown)
        self.Bind(wx.EVT_LEFT_UP, self.OnMouseLeftUp)

        self.Bind(wx.EVT_RIGHT_UP, self.OnMouseRightUp)
        self.history_box.Bind(wx.EVT_CHECKLISTBOX, self.HistoryBoxClicked)
        self.new_nod.Bind(wx.EVT_MOTION, self.NewNodStartDrag)

        self.Bind(wx.EVT_BUTTON, self.ButtonDelete, self.delete_button)
        self.Bind(wx.EVT_BUTTON, self.ButtonAddNode, self.add_button)
        self.Bind(wx.EVT_BUTTON, self.SaveCanvasToDict, self.save_button)
        self.Bind(wx.EVT_BUTTON, self.RecoveryCanvasFromDict, self.load_button)
        self.Bind(wx.EVT_BUTTON, self.ClearCanvas, self.clear_button)

        self.SetDropTarget(TextDropTarget(self))

        self.applicationId = "ConnectedBoxes"
        self.ButtonAddNode(wx._core.CommandEvent)
        self.ButtonAddNode(wx._core.CommandEvent)
        self.ButtonAddNode(wx._core.CommandEvent)
        self.ButtonAddNode(wx._core.CommandEvent)
        self.ButtonAddNode(wx._core.CommandEvent)

    def CreateNodeFromDescriptionAtPosition(self, nodeDescription, pos):
        #We should always get json
        try:
            nodeDescriptionDict = json.loads(nodeDescription)
        except:
            logging.warning("Cannot create a node from a provided description")
            logging.debug("Provided node description should be in JSON format")
            logging.debug(nodeDescription)
            return None

        #We should always get APPLICATION_ID field
        if not "APPLICATION_ID" in nodeDescriptionDict:
            logging.warning("Cannot create a node from a provided description")
            logging.debug("Provided node description should contain APPLICATION_ID field")
            logging.debug(nodeDescription)
            return None

        #Only currently selected APPLICATION_ID is supported
        if nodeDescriptionDict["APPLICATION_ID"] != self.applicationId:
            logging.warning("Cannot create a node from a provided description")
            logging.debug("Provided node description APPLICATION_ID field is incompatible with current application")
            logging.debug(nodeDescription)
            return None

        node = self._nodesFactory.CreateNodeFromDescription(nodeDescriptionDict)
        if node:
            node.position = pos
            self._canvasObjects.append(node)
            self.Render()
            return node
        else:
            logging.warning("Cannot create a node from a provided description")
            logging.debug(nodeDescription)
            return None

    def Render(self):
        """Render all nodes and their connection in depth order."""
        cdc = wx.ClientDC(self)
        self.PrepareDC(cdc)
        dc = wx.BufferedDC(cdc, self._dcBuffer)
        dc.Clear()
        gc = wx.GraphicsContext.Create(dc)

        gc.SetFont(wx.SystemSettings.GetFont(wx.SYS_DEFAULT_GUI_FONT))

        for obj in self._canvasObjects:
            gc.PushState()
            obj.Render(gc)
            gc.PopState()

        if self._objectUnderCursor:
            gc.PushState()
            self._objectUnderCursor.RenderHighlighting(gc)
            if isinstance(self._objectUnderCursor, MoveMe.Canvas.Objects.SimpleScalableTextBoxNode.SimpleScalableTextBoxNode):
                for connection in range(len(self._objectUnderCursor.GetOutcomingConnections())):
                    self._objectUnderCursor.GetOutcomingConnections()[connection].RenderHighlightingIn(gc)

                    for row in range(Matrix.GetNumberRows(self.sheet1) + 1):
                        for col in range(Matrix.GetNumberCols(self.sheet1) + 1):
                            if self._objectUnderCursor.text == unicode(Matrix.GetColLabelValue(self.sheet1, row)) and \
                               self._objectUnderCursor.GetOutcomingConnections()[connection].destination.text == unicode(Matrix.GetColLabelValue(self.sheet1, col)):
                                Matrix.SetCellValue(self.sheet1, row, col, '1')
                                Matrix.SetCellTextColour(self.sheet1, row, col, '#ff0000')

                for connection in range(len(self._objectUnderCursor.GetIncomingConnections())):
                    self._objectUnderCursor.GetIncomingConnections()[connection].RenderHighlightingOut(gc)

                    for row in range(Matrix.GetNumberRows(self.sheet1) + 1):
                        for col in range(Matrix.GetNumberCols(self.sheet1) + 1):
                            if self._objectUnderCursor.text == unicode(Matrix.GetColLabelValue(self.sheet1, row)) and \
                               self._objectUnderCursor.GetIncomingConnections()[connection].source.text == unicode(Matrix.GetColLabelValue(self.sheet1, col)):
                                Matrix.SetCellValue(self.sheet1, col, row, '1')
                                Matrix.SetCellTextColour(self.sheet1, col, row, '#0000ff')

            gc.PopState()

        if not isinstance(self._objectUnderCursor, MoveMe.Canvas.Objects.SimpleScalableTextBoxNode.SimpleScalableTextBoxNode):
            for row in range(Matrix.GetNumberRows(self.sheet1)):
                for col in range(Matrix.GetNumberCols(self.sheet1)):
                    if Matrix.GetCellTextColour(self.sheet1, row, col) == '#ff0000':
                        Matrix.SetCellValue(self.sheet1, row, col, '1')
                        Matrix.SetCellTextColour(self.sheet1, row, col, '#000000')
                    elif Matrix.GetCellTextColour(self.sheet1, row, col) == '#0000ff':
                        Matrix.SetCellValue(self.sheet1, row, col, '1')
                        Matrix.SetCellTextColour(self.sheet1, row, col, '#000000')
#
        if self._selectedObject:
            gc.PushState()
            self._selectedObject.RenderSelection(gc)
            gc.PopState()

        if self._connectionStartObject:
            gc.PushState()
            if self._objectUnderCursor and self._objectUnderCursor.connectableDestination:
                if (self.tree and self._objectUnderCursor.position[1] > self._connectionStartObject.position[1]
                    and (self._objectUnderCursor.position[1] - self._connectionStartObject.position[1]) == 50
                    and self._objectUnderCursor.GetIncomingConnections() == []):
                    gc.SetPen(wx.Pen('#ff0000', 5, wx.DOT_DASH))
                    gc.DrawLines([self._connectionStartObject.GetCenter(), self._objectUnderCursor.GetCenter()])
                elif not self.tree:
                    gc.SetPen(wx.Pen('#ff0000', 5, wx.DOT_DASH))
                    gc.DrawLines([self._connectionStartObject.GetCenter(), self._objectUnderCursor.GetCenter()])
                else:
                    gc.SetPen(wx.Pen('#ccccff', 5, wx.DOT_DASH))
                    gc.DrawLines([self._connectionStartObject.GetCenter(), self._curMousePos])
            else:
                gc.SetPen(wx.Pen('#ccccff', 5, wx.DOT_DASH))
                gc.DrawLines([self._connectionStartObject.GetCenter(), self._curMousePos])
            gc.PopState()

    def OnMouseMotion(self, evt):
        pos = self.CalcUnscrolledPosition(evt.GetPosition()).Get()
        self._objectUnderCursor = self.FindObjectUnderPoint(pos)
        self._curMousePos = pos

        if not evt.LeftIsDown():
            self._draggingObject = None
            self._connectionStartObject = None

        if evt.LeftIsDown() and evt.Dragging() and self._draggingObject:
            xy = self.TreeXY(self._lastDraggingPosition)
            #"липке" переміщення
            newX = self._draggingObject.position[0]+pos[0]-xy[0]
            newY = self._draggingObject.position[1]+pos[1]-xy[1]

            #Check canvas boundaries 
            newX = min(newX, self.canvasDimensions[0]-self._draggingObject.boundingBoxDimensions[0])
            newY = min(newY, self.canvasDimensions[1]-self._draggingObject.boundingBoxDimensions[1])
            newX = max(newX, 0)
            newY = max(newY, 0)

            xy = self.TreeXY([newX, newY])
            self._draggingObject.position = xy

            #видаляємо однорівневі з'єднання та зв'язок, якщо нона надто низько/високо
            if self.tree:
                for connect in self._draggingObject.GetOutcomingConnections():
                    if (connect.destination.position[1] == self._draggingObject.position[1]
                        or (connect.destination.position[1] - self._draggingObject.position[1]) != 50):
                        self._selectedObject = connect
                        self.ButtonDelete(None)
                for connect in self._draggingObject.GetIncomingConnections():
                    if (connect.source.position[1] == self._draggingObject.position[1]
                        or (self._draggingObject.position[1] - connect.source.position[1]) != 50):
                        self._selectedObject = connect
                        self.ButtonDelete(None)

            #Cursor will be at a border of a node if it goes out of canvas
            self._lastDraggingPosition = [min(pos[0], self.canvasDimensions[0]), min(pos[1], self.canvasDimensions[1])]

        self.Render()

    def OnMouseLeftDown(self, evt):
        if not self._objectUnderCursor:
            return

        if evt.ShiftDown():
            if self._objectUnderCursor.connectableSource:
                self._connectionStartObject = self._objectUnderCursor
        elif evt.ControlDown() and self._objectUnderCursor.clonable:
            nodeDescriptionDict = self._objectUnderCursor.GetCloningNodeDescription()
            nodeDescriptionDict["APPLICATION_ID"] = self.applicationId
            data = wx.TextDataObject(json.dumps(nodeDescriptionDict))
            dropSource = wx.DropSource(self)
            dropSource.SetData(data)
            dropSource.DoDragDrop(wx.Drag_AllowMove)
        else:
            if self._objectUnderCursor.movable:
                self._lastDraggingPosition = self.CalcUnscrolledPosition(evt.GetPosition()).Get()
                self._draggingObject = self._objectUnderCursor

        self._lastLeftDownPos = evt.GetPosition()

        self.Render()

    def OnMouseLeftUp(self, evt):
        if (self._connectionStartObject
            and self._objectUnderCursor
            and self._connectionStartObject != self._objectUnderCursor
            and self._objectUnderCursor.connectableDestination):
            if (self.tree and self._objectUnderCursor.position[1] > self._connectionStartObject.position[1]
                and (self._objectUnderCursor.position[1] - self._connectionStartObject.position[1]) == 50
                and self._objectUnderCursor.GetIncomingConnections() == []):
                self.ConnectNodes(self._connectionStartObject, self._objectUnderCursor)
            elif not self.tree:
                self.ConnectNodes(self._connectionStartObject, self._objectUnderCursor)

        #Selection
        if (self._lastLeftDownPos
                and self._lastLeftDownPos[0] == evt.GetPosition()[0]
                and self._lastLeftDownPos[1] == evt.GetPosition()[1]
                and self._objectUnderCursor
                and self._objectUnderCursor.selectable):
            self._selectedObject = self._objectUnderCursor

        self._connectionStartObject = None
        self._draggingObject = None
        self.Render()

    def OnMouseRightUp(self, evt):
        if isinstance(self._objectUnderCursor, MoveMe.Canvas.Objects.Connection.Connection):
            menu = wx.Menu()

            #Append canvas menu items here
            item = wx.MenuItem(menu, wx.NewId(), u"Delete connection")
            menu.AppendItem(item)

            menu.Bind(wx.EVT_MENU, self.UnderRightClickDelete, item)

            #Append node menu items
            if self._objectUnderCursor:
                self._objectUnderCursor.AppendContextMenuItems(menu)

            self.PopupMenu(menu, evt.GetPosition())
            menu.Destroy()

        elif isinstance(self._objectUnderCursor, MoveMe.Canvas.Objects.SimpleScalableTextBoxNode.SimpleScalableTextBoxNode):
            menu = wx.Menu()
            #Append canvas menu items here
            item = wx.MenuItem(menu, 17, u"Delete node whis save connection")
            menu.AppendItem(item)
            item2 = wx.MenuItem(menu, wx.NewId(), u"Delete node")
            menu.AppendItem(item2)

            menu.Bind(wx.EVT_MENU, self.UnderRightClickDelete, item)
            menu.Bind(wx.EVT_MENU, self.UnderRightClickDelete, item2)

            #Append node menu items
            if self._objectUnderCursor:
                self._objectUnderCursor.AppendContextMenuItems(menu)

            self.PopupMenu(menu, evt.GetPosition())
            menu.Destroy()

    def FindObjectUnderPoint(self, pos):
        #Check all objects on a canvas. 
        #Some objects may have multiple components and connections.
        for obj in reversed(self._canvasObjects):
            objUnderCursor = obj.ReturnObjectUnderCursor(pos)
            if objUnderCursor:
                return objUnderCursor
        return None

    def NewNodStartDrag(self, e):
        if e.Dragging():
            data = wx.URLDataObject()

            dropSource = wx.DropSource(self.new_nod)
            dropSource.SetData(data)
            dropSource.DoDragDrop()

    def TreeXY(self, xy):
        def are1stNode():
            are1st = 0
            for node in self._canvasObjects:
                if node.position[1] == 0:
                    if self._draggingObject == node:
                        are1st = 0
                    else:
                        are1st = 1
            if not are1st:
                xy[1] = 0
            else:
                xy[1] = 50
            return xy[1]

        if self.tree:
            xy = list(xy)
            if xy[1] > 367: #обмеження нижньої межі + розмір нода
                xy[1] = 350
                return xy

            if len(list(str(xy[1]))) == 3:
                if int(list(str(xy[1]))[1]) > 5:
                    xy[1] = int(list(str(xy[1]))[0] + '5' + '0')
                else:
                    xy[1] = int(list(str(xy[1]))[0] + '0' + '0')
            elif len(list(str(xy[1]))) == 2:
                if int(list(str(xy[1]))[0]) > 5:
                    xy[1] = 50
                else:
                    xy[1] = are1stNode()
            else:
                xy[1] = are1stNode()

            return xy

        else: #якщо це не дерево
            return xy #повертаємо ху без змін

    def ButtonDelete(self, evt):
        if isinstance(self._selectedObject, MoveMe.Canvas.Objects.SimpleScalableTextBoxNode.SimpleScalableTextBoxNode):
            if self._selectedObject and self._selectedObject.deletable:
                for row in range(Matrix.GetNumberRows(self.sheet1) + 1):
                    for col in range(Matrix.GetNumberCols(self.sheet1) + 1):
                        if self._selectedObject.text == unicode(Matrix.GetColLabelValue(self.sheet1, row)) and self._selectedObject.text == unicode(Matrix.GetColLabelValue(self.sheet1, col)):
                            Matrix.DeleteRows(self.sheet1, row)
                            Matrix.DeleteCols(self.sheet1, col)

                self.numberNods.pop(self.numberNods.index(int(self._selectedObject.text)))

                n = 0
                for name in self.numberNods:
                    Matrix.SetColLabelValue(self.sheet1, n, str(name))
                    Matrix.SetRowLabelValue(self.sheet1, n, str(name))
                    n += 1

                self.SaveCanvasToDict('Delete node ' + self._selectedObject.text)
                for InConnection in self._selectedObject.GetIncomingConnections():
                    self.SaveCanvasToDict('Auto delete connection ' + self._selectedObject.text + ' -> ' + InConnection.source.text)
                for OutConnection in self._selectedObject.GetOutcomingConnections():
                    self.SaveCanvasToDict('Auto delete connection ' + self._selectedObject.text + ' -> ' + OutConnection.destination.text)

                self._selectedObject.Delete()
                if self._selectedObject in self._canvasObjects:
                    self._canvasObjects.remove(self._selectedObject)
                self._selectedObject = None
                self.Render()
        elif isinstance(self._selectedObject, MoveMe.Canvas.Objects.Connection.Connection):
            if self._selectedObject and self._selectedObject.deletable:
                for row in range(Matrix.GetNumberRows(self.sheet1) + 1):
                    for col in range(Matrix.GetNumberCols(self.sheet1) + 1):
                        if self._selectedObject.source.text == unicode(Matrix.GetColLabelValue(self.sheet1, row)) and self._selectedObject.destination.text == unicode(Matrix.GetColLabelValue(self.sheet1, col)):
                            Matrix.SetCellValue(self.sheet1, row, col, '')

                self.SaveCanvasToDict('Delete connection ' + self._selectedObject.source.text + ' -> ' + self._selectedObject.destination.text)
                self._selectedObject.Delete()
                if self._selectedObject in self._canvasObjects:
                    self._canvasObjects.remove(self._selectedObject)
                self.__allConectionsNumber -= 1
                self._selectedObject = None
                self.Render()

    def ButtonAddNode(self, e, xy=[]):
        if self.numberNods: #якщо список нод не пустий
            list_number_nods = range(1, self.numberNods[-1] + 1)
            for node in list_number_nods:
                if node not in self.numberNods:
                    self.numberForNod = node
                    break
                else:
                    self.numberForNod = self.numberNods[-1] + 1
                    while 1:
                        if self.numberForNod in self.numberNods:
                            self.numberForNod += 1
                        else:
                            break

        if not xy:
            xy = [randint(0, 560), randint(0, 370)]

        xy = self.TreeXY(xy)

        self.CreateNodeFromDescriptionAtPosition('{"NodeClass": "SimpleScalableTextBoxNode", "APPLICATION_ID": "ConnectedBoxes", "NodeParameters":{"text":"'
                                                 + str(self.numberForNod) + '"}}', (xy[0], xy[1]))
        self.numberNods.append(self.numberForNod)
        Matrix.SetNumberCols(self.sheet1, Matrix.GetNumberCols(self.sheet1) + 1)
        Matrix.SetNumberRows(self.sheet1, Matrix.GetNumberRows(self.sheet1) + 1)

        self.numberNods.sort() #щоб був впорядкований список
        n = 0
        for name in self.numberNods:
            Matrix.SetColLabelValue(self.sheet1, n, str(name))
            Matrix.SetRowLabelValue(self.sheet1, n, str(name))
            n += 1

        self.SaveCanvasToDict('Add node '+ str(self.numberForNod))

    def UnderRightClickDelete(self, evt):
        if isinstance(self._objectUnderCursor, MoveMe.Canvas.Objects.SimpleScalableTextBoxNode.SimpleScalableTextBoxNode):
            if self._objectUnderCursor and self._objectUnderCursor.deletable:
                Matrix.SetCellValue(self.sheet1, int(self._objectUnderCursor.text) - 1, int(self._objectUnderCursor.text) - 1, '')
                Matrix.DeleteRows(self.sheet1, int(self._objectUnderCursor.text) - 1)
                Matrix.DeleteCols(self.sheet1, int(self._objectUnderCursor.text) - 1)

                self.numberNods.pop(self.numberNods.index(int(self._objectUnderCursor.text)))

                n = 0
                for name in self.numberNods:
                    Matrix.SetColLabelValue(self.sheet1, n, str(name))
                    Matrix.SetRowLabelValue(self.sheet1, n, str(name))
                    n += 1

                #delete whis save
                if evt.GetEventObject().GetLabelText(17):
                    for InConnection in self._objectUnderCursor.GetIncomingConnections():
                        self.SaveCanvasToDict('Auto delete connection ' + InConnection.source.text + ' -> ' + InConnection.destination.text)
                    for OutConnection in self._objectUnderCursor.GetOutcomingConnections():
                        self.SaveCanvasToDict('Auto delete connection ' + OutConnection.source.text + ' -> ' + OutConnection.destination.text)
                    self.__allConectionsNumber -= 1
                    self.SaveCanvasToDict('Delete node ' + self._objectUnderCursor.text + ' whis save')
                else:
                    self.SaveCanvasToDict('Delete node ' + self._objectUnderCursor.text)

                self._objectUnderCursor.Delete()
                if self._objectUnderCursor in self._canvasObjects:
                    self._canvasObjects.remove(self._objectUnderCursor)

                for connectIN in self._objectUnderCursor._incomingConnections:
                    for connectOUT in self._objectUnderCursor._outcomingConnections:
                        self.ConnectNodes(connectIN.source, connectOUT.destination, part_text='Auto a')

                self.__allConectionsNumber -= 1
                self._selectedObject = None
                self._objectUnderCursor = None
                self.Render()

    def OnKeyPress(self, evt):
        if evt.GetKeyCode() == wx.WXK_SHIFT: #ignore shift key (for create connect)
            return
        if evt.GetKeyCode() == wx.WXK_CONTROL: #ignore ctrl key (for clone connect)
            return
        if evt.GetKeyCode() == wx.WXK_DELETE:
            if self._selectedObject and self._selectedObject.deletable:
                self._selectedObject.Delete()
                if self._selectedObject in self._canvasObjects:
                    self._canvasObjects.remove(self._selectedObject)
                self._selectedObject = None
        elif evt.GetKeyCode() == wx.WXK_RETURN:
            MatrixToRecovery = self.History[-1]['Matrix']
            for row in range(Matrix.GetNumberRows(MatrixToRecovery)):
                for col in range(Matrix.GetNumberCols(MatrixToRecovery)):
                    if Matrix.GetCellValue(self.sheet1, row, col) != Matrix.GetCellValue(MatrixToRecovery, row, col):
                        if Matrix.GetCellValue(self.sheet1, row, col) == u'1':
                            if self.tree and self._canvasObjects[row].position[1] < self._canvasObjects[col].position[1]:
                                self.ConnectNodesByIndexes(row, col)
                            elif not self.tree:
                                self.ConnectNodesByIndexes(row, col)
                            else:
                                Matrix.SetCellValue(self.sheet1, row, col, '')
                        elif Matrix.GetCellValue(self.sheet1, row, col) == u'':
                            for connectOUT in self._canvasObjects[row].GetOutcomingConnections(): #забороняємо декілька стрілок із однаковим напрямком
                                for connectIN in self._canvasObjects[col].GetIncomingConnections():
                                    if connectIN == connectOUT:
                                        self._selectedObject = connectIN
                                        self.ButtonDelete(None)
                                        break
                        else:
                            Matrix.SetCellValue(self.sheet1, row, col, '')
        else:
            evt.Skip()

        #Update object under cursor                
        pos = self.CalcUnscrolledPosition(evt.GetPosition()).Get()
        self._objectUnderCursor = self.FindObjectUnderPoint(pos)

        self.Render()

    def ConnectNodes(self, source, destination, part_text='A'):
        for connectIN in destination.GetIncomingConnections(): #забороняємо декілька стрілок із однаковим напрямком
            for connectOUT in source.GetOutcomingConnections():
                if connectIN == connectOUT:
                    return

        newConnection = self._connectionsFactory.CreateConnectionBetweenNodesFromDescription(source, destination)
        if newConnection:
            source.AddOutcomingConnection(newConnection)
            destination.AddIncomingConnection(newConnection)

            for row in range(Matrix.GetNumberRows(self.sheet1)):
                for col in range(Matrix.GetNumberCols(self.sheet1)):
                    if source.text == unicode(Matrix.GetColLabelValue(self.sheet1, row)) and destination.text == unicode(Matrix.GetColLabelValue(self.sheet1, col)):
                        Matrix.SetCellValue(self.sheet1, row, col, '1')

            self.__allConectionsNumber += 1
            self.SaveCanvasToDict(part_text + 'dd connection ' + source.text + ' -> ' + destination.text)

    def ConnectNodesByIndexes(self, sourceIdx, destinationIdx):
        self.ConnectNodes(self._canvasObjects[sourceIdx], self._canvasObjects[destinationIdx])

    def CurrentTime(self):
        return str(datetime.now()).split(' ')[1].split('.')[0]

    def SaveCanvasToDict(self, e):
        self.nb2 = wx.Notebook(self, -1, (-10, -10), style=wx.NB_HITTEST_NOWHERE)
        MatrixToSave = Matrix(self.nb2)
        CanvasToSave = {}

        #Save nodes
        CanvasToSave["Nodes"] = []
        for node in self._canvasObjects:
            CanvasToSave["Nodes"].append(node.SaveObjectToDict())

        #Save connections
        CanvasToSave["Connections"] = []
        for node in self._canvasObjects:
            for connection in node.GetOutcomingConnections():
                CanvasToSave["Connections"].append({"sourceIdx":self._canvasObjects.index(connection.source), "destinationIdx":self._canvasObjects.index(connection.destination)})

        #self.SavedMatrix = self.sheet1 create link to object; copy.deepcopy don't clone generators; becouse we use
        CanvasToSave["Matrix"] = 0
        Matrix.SetNumberCols(MatrixToSave, Matrix.GetNumberCols(self.sheet1))
        Matrix.SetNumberRows(MatrixToSave, Matrix.GetNumberRows(self.sheet1))
        for row in range(Matrix.GetNumberRows(self.sheet1)):
            for col in range(Matrix.GetNumberCols(self.sheet1)):
                Matrix.SetRowLabelValue(MatrixToSave, row, Matrix.GetRowLabelValue(self.sheet1, row))
                Matrix.SetColLabelValue(MatrixToSave, col, Matrix.GetColLabelValue(self.sheet1, col))
                Matrix.SetCellValue(MatrixToSave, row, col, Matrix.GetCellValue(self.sheet1, row, col))
        CanvasToSave["Matrix"] = MatrixToSave

        if isinstance(e, wx._core.CommandEvent):
            self.AllSavedCanvas.append(CanvasToSave)
            self.saved_canvas_box.Append(self.CurrentTime() + ' ' + 'Nodes: ' + str(len(self._canvasObjects)) + ' Connections: ' + str(self.__allConectionsNumber))
            self.AllSavedHistory.append(self.history_box.GetItems())
        else:
            if self.history_box.GetChecked()  and 'connection' in self.history_box.GetItems()[self.history_box.GetChecked()[0]]:
                return

            if self.ItemsToDelete: #якщо список об"єктів для видалення не пустий
                if len(self.ItemsToDelete) == 1:
                    self.history_box.SetItems(self.history_box.GetItems()[:self.ItemsToDelete[-1]])
                    self.History = self.History[:self.ItemsToDelete[-1]]
                else:
                    self.history_box.SetItems(self.history_box.GetItems()[:self.ItemsToDelete[0]]) #видаляємо все після відміченого
                    self.History = self.History[:self.ItemsToDelete[0]]
                self.ItemsToDelete = []

            self.History.append(CanvasToSave)
            self.history_box.Append(self.CurrentTime() + ' ' + e)
            self.history_box.LineDown() #прокручуємо історію вниз

    def RecoveryCanvasFromDict(self, e):
        if isinstance(e, wx._core.CommandEvent):
            self.history_box.Clear()
            if len(self.saved_canvas_box.GetChecked()) != 1:
                for i in self.saved_canvas_box.GetChecked():
                    self.saved_canvas_box.Check(i, 0)
                return
            SavedCanvasToRecovery = self.AllSavedCanvas[self.saved_canvas_box.GetChecked()[0]]
            if not self.history_box.GetItems(): #щоб не заповнювало історію вдруге
                for item in self.AllSavedHistory[self.saved_canvas_box.GetChecked()[0]]:
                    if 'connection' not in item:
                        self.history_box.Append(item)

        else:
            SavedCanvasToRecovery = self.History[self.history_box.GetChecked()[0]]

        self.ClearCanvas(None)

        #Recovery nodes
        for i in range(len(SavedCanvasToRecovery["Nodes"])):
            if SavedCanvasToRecovery["Nodes"][i]['NodeParameters']['_outcomingConnections'] or SavedCanvasToRecovery["Nodes"][i]['NodeParameters']['_incomingConnections']:
                SavedCanvasToRecovery["Nodes"][i]['NodeParameters']['_outcomingConnections'] = []
                SavedCanvasToRecovery["Nodes"][i]['NodeParameters']['_incomingConnections'] = []
        for nodeDict in SavedCanvasToRecovery["Nodes"]:
            newNode = self._nodesFactory.CreateNodeFromDescription(nodeDict)
            if newNode:
                self._canvasObjects.append(newNode)
            else:
                logging.error("Cannot recovery saved node")

        #Recovery connections
        for connectionDict in SavedCanvasToRecovery["Connections"]:
            if isinstance(connectionDict["sourceIdx"], int) and isinstance(connectionDict["destinationIdx"], int):
                self.ConnectNodesByIndexes(connectionDict["sourceIdx"], connectionDict["destinationIdx"])

        #Recovery matrix
        MatrixToRecovery = SavedCanvasToRecovery["Matrix"]
        Matrix.SetNumberCols(self.sheet1, Matrix.GetNumberCols(MatrixToRecovery))
        Matrix.SetNumberRows(self.sheet1, Matrix.GetNumberRows(MatrixToRecovery))
        for row in range(Matrix.GetNumberRows(MatrixToRecovery)):
            for col in range(Matrix.GetNumberCols(MatrixToRecovery)):
                Matrix.SetRowLabelValue(self.sheet1, row, Matrix.GetRowLabelValue(MatrixToRecovery, row))
                Matrix.SetColLabelValue(self.sheet1, col, Matrix.GetColLabelValue(MatrixToRecovery, col))
                Matrix.SetCellValue(self.sheet1, row, col, Matrix.GetCellValue(MatrixToRecovery, row, col))
                if int(Matrix.GetRowLabelValue(MatrixToRecovery, row)) not in self.numberNods: #якщо рядка немає у списку нод
                    self.numberNods.append(int(Matrix.GetRowLabelValue(MatrixToRecovery, row))) #додаємо

        for i in self.saved_canvas_box.GetChecked():
            self.saved_canvas_box.Check(i, 0)

        self.Render()

    def HistoryBoxClicked(self, e):
        self.ItemsToDelete = []

        if len(self.history_box.GetChecked()) == 1:
            for i in range(len(self.history_box.GetItems())):
                self.history_box.SetItemBackgroundColour(i,  '#FFFFFF')
                self.history_box.SetSelection(i) #вирішення для багу у wx із оновленням
            self.history_box.SetSelection(-1) #вирішення для багу у wx із оновленням
            if 'save' in self.history_box.GetItems()[self.history_box.GetChecked()[0]]:
                for i in self.history_box.GetChecked():
                    self.history_box.Check(i, 0)
                return

        for i in range(self.history_box.GetChecked()[0] + 1, len(self.history_box.GetItems())):
            self.ItemsToDelete.append(i)
            self.history_box.SetItemBackgroundColour(i,  '#cecece')
            self.history_box.SetSelection(i) #вирішення для багу у wx із оновленням
        self.history_box.SetSelection(-1) #вирішення для багу у wx із оновленням

        self.RecoveryCanvasFromDict(None)

        for i in self.history_box.GetChecked():
            self.history_box.Check(i, 0)

    def ClearCanvas(self, e):
        for node in self._canvasObjects:
            node.Delete()
            node.DeleteAllOutcomingConnections()
            node.DeleteAllIncomingConnections()
        Matrix.SetNumberCols(self.sheet1, 0)
        Matrix.SetNumberRows(self.sheet1, 0)
        self.numberNods = []
        self.numberForNod = 1
        self._canvasObjects = []
        self.__allConectionsNumber = 0
        if e != None:
            self.history_box.Clear()
        self.Render()