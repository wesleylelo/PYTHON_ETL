import socket

def send_question (client_socket, question):
    
    client_socket.send(question.encode())


def receive_answer (client_socket):
    
    return client_socket.recv(1024).decode().strip()

def main ():
    
    host = '127.0.0.2'

    port = 15001
    
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    
    server_socket.bind((host, port))
    
    server_socket.listen(5)
    
    server_socket.settimeout(50.0)

    print ("Servidor está esperando conexões...")

    client_socket, addr = server_socket.accept()
    
    play_addr = 0
    
    print ( "Cliente conectado:" , addr)
    
    while True :
        for x in range(3):
            
            question = ["Qual é a capital de Rondônia Brasil?","Quem perdeu o primeiro pênalti entre Brasil e Croácia na copa de 2022?","Quem é o melhor batedor de Pênalte do Mundo?"]

            send_question(client_socket, question[x])
            
            answers = ["Porto Velho", "Rodrigo", "Neymar"]

            answer = receive_answer(client_socket)     
            
            print ("Resposta do cliente:" , answer)
            
            if answer == answers[x]:
               play_addr += 1
        
        print("Jogador ", addr, " ficou com ", play_addr, " pontos")
        
        terminal = "Você quer continuar? Digite 1, para sim, e 0 para não"
        
        send_question(client_socket,terminal)
        
        answer = receive_answer(client_socket)
        
        ler = int(answer)
        
        if ler == 0:
           
           client_socket.close()
           
           break

if __name__ == "__main__":
    main()
