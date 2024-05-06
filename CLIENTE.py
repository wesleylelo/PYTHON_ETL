import  socket

def receive_question (server_socket):
    
    return server_socket.recv(1024).decode().strip()

def send_answer (server_socket, answer):
    
    server_socket.send(answer.encode())

def main ():
    
    host = '127.0.0.2'

    port = 15001
    
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    
    client_socket.connect((host,port))
    
    answer = 1

    while(answer != 0):

        question = receive_question(client_socket)
    
        print ("Pergunta:" , question)

        answer = input ("Sua resposta: ")

        send_answer(client_socket, answer)

if __name__ == "__main__":
    main()
