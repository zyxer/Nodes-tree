import wx
from MoveMe.Canvas.Canvas import Canvas
from MoveMe.Canvas.Factories.DefaultNodesFactory import DefaultNodesFactory
from MoveMe.Canvas.Objects.SimpleScalableTextBoxNode import SimpleScalableTextBoxNode

class CanvasWindow(wx.Frame):
    def __init__(self,  *args, **kw):
        wx.Frame.__init__(self, size=(1000, 600), style=(wx.CAPTION | wx.MINIMIZE_BOX | wx.CLOSE_BOX | wx.SYSTEM_MENU), *args, **kw)
        s = wx.BoxSizer(wx.VERTICAL)


        supportedClasses = {"SimpleScalableTextBoxNode":SimpleScalableTextBoxNode}

        tree = 1
        canvas = Canvas(parent=self, tree=tree, nodesFactory=DefaultNodesFactory(supportedClasses))

        self.Bind(wx.EVT_CHAR_HOOK, canvas.OnKeyPress)

        s.Add(canvas, 1, wx.EXPAND)
        self.SetSizer(s)
        if tree:
            self.SetTitle("Processing tree graph models")
        else:
            self.SetTitle("Processing graph models")

if __name__ == '__main__':
    app = wx.PySimpleApp()
    CanvasWindow(None).Show()
    app.MainLoop()