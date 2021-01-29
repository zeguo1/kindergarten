FROM registry.cn-qingdao.aliyuncs.com/qy-dockerhub/python:3.6.2

#RUN cp /usr/share/zoneinfo/Asia/Shanghai /etc/localtime && echo 'Asia/Shanghai'>/etc/timezone

ENV FLASK_ENV=production

WORKDIR /root/kindergarten/

RUN pip install --upgrade pip -i http://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
COPY . .

RUN pip install -i http://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com --no-cache-dir -r requirements.txt

EXPOSE 5000
CMD [ "flask", "run" ,"--host=0.0.0.0" ]