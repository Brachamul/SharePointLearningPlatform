import SimpleHTTPServer
import SocketServer

Handler = SimpleHTTPServer.SimpleHTTPRequestHandler
server = SocketServer.TCPServer(('0.0.0.0', 8000), Handler)

server.serve_forever()