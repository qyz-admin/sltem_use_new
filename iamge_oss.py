import oss2, uuid

class AliyunOss():
    def __init__(self):
        self.access_key_id = "[AccessKey ID]"
        self.access_key_secret = "[Secret]"
        self.auth = oss2.Auth(self.access_key_id, self.access_key_secret)
        self.bucket_name = "[doublez-mytest]"
        self.endpoint = "[oss-cn-shanghai.aliyuncs.com]"
        self.bucket = oss2.Bucket(self.auth, self.endpoint, self.bucket_name)

    def put_object_from_file(self, name, file):
        self.bucket.put_object_from_file(name, file)
        return "https://{}.{}/{}".format(self.bucket_name, self.endpoint, name)





if __name__ == '__main__':
    im = AliyunOss()
    img_url = im.put_object_from_file("target_name.png", "img.png")
