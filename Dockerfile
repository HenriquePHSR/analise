FROM python:3.8.3-slim
LABEL maintainer="Pedro Rodrigues <pir.pedro@gmail.com>"

RUN apt-get update && \
    apt-get install --no-install-recommends -y sudo && \
    apt-get clean && rm -rf /var/lib/apt/lists/* 

RUN groupadd -g 1000 app \
    && useradd -u 1000 -g app -s /bin/bash -m app \
    && usermod -aG sudo app \
    && echo 'app    ALL=(ALL) NOPASSWD:ALL' >> /etc/sudoers \
    # && mkdir -p /usr/src/app \
    && chown -R app:app /usr/src 
# && chmod -R 777 /usr/src/app

RUN pip3 install --upgrade pip \
    && pip3 install numpy scipy pandas matplotlib sklearn imblearn jupyter

# WORKDIR /usr/src/app
