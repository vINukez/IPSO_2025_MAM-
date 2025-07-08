from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_code_paragraph(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    font = run.font
    font.name = 'Courier New'
    font.size = Pt(10)
    return p

# Neues Dokument erstellen
doc = Document()

# Titel
doc.add_heading('Umfassender Leitfaden zu Nmap und dessen Parametern', level=1)

# Einleitung
doc.add_heading('Einleitung', level=2)
doc.add_paragraph(
    "Nmap (Network Mapper) ist ein kostenloses und Open-Source-Netzwerksicherheitstool, "
    "das von Gordon Lyon (alias Fyodor) entwickelt wurde. Es wird verwendet, um Netzwerkverbindungen "
    "zu scannen und Sicherheitslücken zu identifizieren. Nmap kann verwendet werden, um Hosts, Dienste, "
    "Betriebssysteme, Packet-Filter/Firewalls und andere Eigenschaften eines Netzwerks zu erkennen. Es ist "
    "ein unverzichtbares Werkzeug für Netzwerkadministratoren, Sicherheitsexperten und Penetrationstester."
)

# Installation von Nmap
doc.add_heading('Installation von Nmap', level=2)

doc.add_heading('Auf Linux:', level=3)
add_code_paragraph(doc, "sudo apt-get install nmap")

doc.add_heading('Auf macOS:', level=3)
add_code_paragraph(doc, "brew install nmap")

doc.add_heading('Auf Windows:', level=3)
doc.add_paragraph(
    "Laden Sie das Installationsprogramm von der offiziellen "
    "[Nmap-Website](https://nmap.org/download.html) herunter und folgen Sie den Installationsanweisungen."
)

# Grundlegende Verwendung
doc.add_heading('Grundlegende Verwendung', level=2)
doc.add_paragraph(
    "Der grundlegende Befehl für einen Nmap-Scan lautet:"
)
add_code_paragraph(doc, "nmap [Optionen] [Ziel]")
doc.add_paragraph("Beispiel:")
add_code_paragraph(doc, "nmap scanme.nmap.org")

# Wichtige Nmap-Parameter
doc.add_heading('Wichtige Nmap-Parameter', level=2)

parameters = [
    ("-sS (TCP SYN-Scan)", "Der TCP SYN-Scan ist der Standard-Scanmodus und einer der schnellsten und unauffälligsten Scans. Er sendet ein SYN-Paket an den Zielport und wartet auf eine Antwort.", "nmap -sS [Ziel]"),
    ("-sT (TCP Connect-Scan)", "Dieser Scan verwendet das normale Verbindungsaufbauverfahren des Betriebssystems. Es wird ein vollständiger TCP-Handshake durchgeführt.", "nmap -sT [Ziel]"),
    ("-sU (UDP-Scan)", "Scannt nach offenen UDP-Ports, da UDP-Verbindungen verbindungslos sind. Dies kann langsamer sein als ein TCP-Scan.", "nmap -sU [Ziel]"),
    ("-p (Ports)", "Gibt die zu scannenden Ports an. Standardmäßig scannt Nmap die 1000 am häufigsten verwendeten Ports. Sie können auch bestimmte Ports oder Portbereiche angeben.", "nmap -p 80,443 [Ziel]\nnmap -p 1-65535 [Ziel]"),
    ("-A (Erweiterte Erkennung)", "Aktiviert erweiterte Erkennung, einschließlich Betriebssystem- und Dienstversionserkennung, sowie Skript-Scanning.", "nmap -A [Ziel]"),
    ("-O (Betriebssystem-Erkennung)", "Versucht, das Betriebssystem des Ziels anhand von TCP/IP-Stack-Eigenschaften zu erkennen.", "nmap -O [Ziel]"),
    ("-v (Verbosity)", "Erhöht die Ausführlichkeit der Ausgabe, was nützlich ist, um detaillierte Informationen während des Scans zu erhalten.", "nmap -v [Ziel]"),
    ("-Pn (Ping Scan)", "Deaktiviert den Host-Erreichbarkeitscheck und scannt das Ziel direkt, was nützlich ist, wenn Firewalls ICMP-Pakete blockieren.", "nmap -Pn [Ziel]"),
    ("-sV (Versionserkennung)", "Erkennt die Version der Dienste, die auf den offenen Ports laufen, indem spezielle Pakete gesendet werden.", "nmap -sV [Ziel]"),
    ("--script", "Führt Nmap Scripting Engine (NSE) Skripte aus, um detailliertere Informationen zu sammeln. Es gibt viele Skripte, die verschiedene Arten von Informationen sammeln können.", "nmap --script=[Script] [Ziel]\nnmap --script=http-title [Ziel]"),
    ("-oN (Normal Output)", "Speichert die Scan-Ergebnisse in einer Textdatei im normalen Ausgabeformat.", "nmap -oN output.txt [Ziel]"),
    ("-oX (XML Output)", "Speichert die Scan-Ergebnisse in einer XML-Datei, die maschinenlesbar ist.", "nmap -oX output.xml [Ziel]"),
    ("-oG (Grepable Output)", "Speichert die Scan-Ergebnisse in einem grepfähigen Format.", "nmap -oG output.gnmap [Ziel]"),
    ("-oA (Alle Formate)", "Speichert die Scan-Ergebnisse in allen drei Formaten: Normal, XML und Grepfähig.", "nmap -oA output [Ziel]"),
    ("-iL (Input from List)", "Liest eine Liste von Zielen aus einer Datei.", "nmap -iL targets.txt"),
    ("-iR (Random Targets)", "Scannt eine zufällige Anzahl von Zielen.", "nmap -iR 10"),
    ("-exclude (Ausschließen von Hosts)", "Schließt bestimmte Hosts oder Netzwerke vom Scan aus.", "nmap [Ziel] --exclude 192.168.1.1"),
    ("-F (Schneller Scan)", "Führt einen schnellen Scan durch, der nur die 100 am häufigsten verwendeten Ports überprüft.", "nmap -F [Ziel]"),
    ("--top-ports", "Scannt die angegebenen Top-N häufigsten Ports.", "nmap --top-ports 20 [Ziel]"),
    ("-T (Timing Optionen)", "Passt die Timing-Einstellungen an. Es gibt sechs Stufen von -T0 (langsam und schonend) bis -T5 (schnell und aggressiv).", "nmap -T4 [Ziel]"),
    ("-D (Decoy)", "Verwendet Decoys, um den Ursprung des Scans zu verschleiern.", "nmap -D RND:10 [Ziel]"),
    ("-S (Spoof Source IP)", "Spooft die Quell-IP-Adresse.", "nmap -S 192.168.1.1 [Ziel]"),
    ("-f (Fragmentierung)", "Fragmentiert Pakete, um Intrusion Detection Systems (IDS) zu umgehen.", "nmap -f [Ziel]"),
    ("-g (Source Port)", "Setzt den Quellport des Pakets.", "nmap -g 53 [Ziel]"),
    ("--spoof-mac", "Spooft die MAC-Adresse.", "nmap --spoof-mac 00:11:22:33:44:55 [Ziel]")
]

for param, desc, cmd in parameters:
    doc.add_heading(param, level=3)
    doc.add_paragraph(desc)
    add_code_paragraph(doc, cmd)

# Beispielscans
doc.add_heading('Beispielscans', level=2)

examples = [
    ("Ein einfacher Port-Scan eines Ziels", "nmap scanme.nmap.org"),
    ("Ein Scan mit Betriebssystem-Erkennung und Dienstversionserkennung", "nmap -A scanme.nmap.org"),
    ("Ein Scan eines bestimmten Ports (z.B. 80 und 443)", "nmap -p 80,443 scanme.nmap.org"),
    ("Ein UDP-Scan eines Ziels", "nmap -sU scanme.nmap.org"),
    ("Ein schneller Scan der Top 20 Ports eines Ziels", "nmap --top-ports 20 scanme.nmap.org"),
    ("Ein Scan mit Decoys", "nmap -D RND:10 scanme.nmap.org"),
    ("Ein Scan mit Spoofing der Quell-IP", "nmap -S 192.168.1.1 scanme.nmap.org")
]

for desc, cmd in examples:
    doc.add_heading(desc, level=3)
    add_code_paragraph(doc, cmd)

# Fazit
doc.add_heading('Fazit', level=2)
doc.add_paragraph(
    "Nmap ist ein leistungsstarkes Tool für Netzwerkscans und Sicherheitsbewertungen. "
    "Mit seinen vielfältigen Parametern und Funktionen kann es an die spezifischen Bedürfnisse "
    "verschiedener Szenarien angepasst werden. Die Kenntnis der grundlegenden Parameter und deren "
    "Verwendung ist unerlässlich für effektive Netzwerksicherheit und Analyse."
)

# Dokument speichern
doc.save("nmap-leitfaden.docx")
