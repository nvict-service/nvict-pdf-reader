# -*- coding: utf-8 -*-
"""Zorgt dat er maar één instance van NVict Reader draait wanneer een PDF
wordt geopend (bv. via "Openen met" of dubbelklikken in de Verkenner) -
poort van de tkinter-editie's SingleInstance-klasse (NVict_Reader.py), die
daar een eigen TCP-socket op 127.0.0.1:52847 gebruikte. Hier via Qt's
QLocalServer/QLocalSocket (named pipe op Windows) i.p.v. een TCP-poort -
geen portconflicten of firewall-meldingen, en werkt netjes samen met de
Qt-event-loop in plaats van een eigen thread nodig te hebben.

Zelfde gedrag als tkinter: alleen wanneer er een bestand wordt meegegeven
én er al een instance draait, wordt dat bestand naar de bestaande instance
gestuurd (als nieuwe tab) en sluit deze tweede poging direct af. Zonder
bestand (bv. de exe zelf starten terwijl de app al open is) start gewoon
een nieuwe, aparte instance - dat kan de gebruiker bewust willen."""

from PySide6.QtCore import QObject, Signal
from PySide6.QtNetwork import QLocalServer, QLocalSocket

_SERVER_NAME = "NVictReaderQt-SingleInstance"


def try_send_to_running_instance(file_path):
    """Probeert `file_path` naar een al draaiende instance te sturen.
    Geeft True terug als dat gelukt is (deze instance moet dan stoppen)."""
    socket = QLocalSocket()
    socket.connectToServer(_SERVER_NAME)
    if not socket.waitForConnected(500):
        return False
    socket.write(file_path.encode("utf-8"))
    socket.waitForBytesWritten(500)
    # Wacht op de bevestiging van de server voordat dit process afsluit -
    # zonder dat kan de named pipe al dicht zijn (proces exit) vóórdat de
    # server de data heeft kunnen lezen, en gaat het bestandspad verloren.
    socket.waitForReadyRead(1000)
    return True


class SingleInstanceServer(QObject):
    """Luistert naar bestandspaden die andere (tweede) instances doorsturen."""

    file_received = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._server = QLocalServer(self)
        self._server.newConnection.connect(self._on_new_connection)

    def start(self):
        # Verwijder een eventueel achtergebleven server (bv. na een crash)
        # voordat we opnieuw gaan luisteren - anders faalt listen() stil.
        QLocalServer.removeServer(_SERVER_NAME)
        self._server.listen(_SERVER_NAME)

    def _on_new_connection(self):
        socket = self._server.nextPendingConnection()
        if socket is None:
            return
        # De data kan al binnen zijn vóór we hier komen (dan geeft
        # waitForReadyRead() False terug omdat er geen NIEUW readyRead-
        # signaal meer komt) - alleen wachten als er nog niets ligt.
        if socket.bytesAvailable() == 0:
            socket.waitForReadyRead(1000)
        data = bytes(socket.readAll()).decode("utf-8")
        if data:
            self.file_received.emit(data)
        # Bevestiging terugsturen zodat de verzendende (tweede) instance
        # weet dat het bestandspad is aangekomen en veilig kan afsluiten.
        socket.write(b"ok")
        socket.waitForBytesWritten(500)
        socket.disconnectFromServer()
