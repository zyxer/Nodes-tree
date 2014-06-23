
from MoveMe.Canvas.Objects.Base.ConnectableObject import ConnectableObject

class ConnectableDestination(ConnectableObject):
    def __init__(self):
        super(ConnectableDestination, self).__init__()
        self.connectableDestination = True
        
        self._incomingConnections = []
        
    def AddIncomingConnection(self, connection):
        self._incomingConnections.append(connection)
        
    def DeleteIncomingConnection(self, connection):
        self._incomingConnections.remove(connection)

    def DeleteAllIncomingConnections(self):
        self._incomingConnections = []

    def GetIncomingConnections(self):
        return self._incomingConnections