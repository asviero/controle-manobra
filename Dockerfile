# Usa uma imagem leve do Node.js
FROM node:18-alpine

# Cria pasta de trabalho dentro do container
WORKDIR /app

# Copia os arquivos de configuração primeiro (para aproveitar cache)
COPY package*.json ./

# Instala as dependências
RUN npm install

# Copia o restante do código
COPY . .

# Expõe a porta 3000
EXPOSE 3000

# Comando para rodar
CMD ["node", "server.js"]