FROM node:slim
WORKDIR /app
COPY . /app 
RUN npm install
EXPOSE 80
CMD [ "npm","start" ]
