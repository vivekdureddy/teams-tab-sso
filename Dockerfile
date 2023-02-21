FROM node:lts-alpine
RUN apk update && \
    apk upgrade
RUN mkdir /app
WORKDIR /app
COPY . /app
RUN npm install
ENV PORT=3978
EXPOSE 3978
CMD ["npm", "start"]