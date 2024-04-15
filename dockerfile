FROM node:20.12.0-alpine3.18

WORKDIR /usr/app

COPY package*.json ./

RUN npm install
COPY . .
EXPOSE 3001
CMD [ "npm", "run", "dev"]