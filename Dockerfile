FROM node:18-alpine

WORKDIR /app

COPY package.json .
RUN npm install

COPY server_cesapre.js .

EXPOSE 3000

CMD ["node", "server_cesapre.js"]
