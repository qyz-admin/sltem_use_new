# -*- coding: utf-8 -*-
# @Time    : 2020/9/12 21:01
# @Author  : XiaYouRan
# @Email   : youran.xia@foxmail.com
# @File    : kugou_music2.py
# @Software: PyCharm
import time
from hashlib import md5
import json
import requests
import re
import os

class KuGouMusic(object):
    def __init__(self):
        self.headers = {'referer':'https://www.kugou.com/','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'}
    def MD5Encrypt(self, keyword):
        # 返回当前时间的时间戳(1970纪元后经过的浮点秒数)
        k = time.time()
        mid = int(round(k * 1000))
        clienttime = int(time.time() * 1000)
        info = ["NVPh5oo715z5DIWAeQlhMDsWXXQV4hwt", "appid=1014", "bitrate=0", "callback=callback123", "clienttime={}".format(clienttime), "clientver=1000", "dfid=4XSQkz1mmSGI2XV1Ud1xgR9V",
                "filter=10", "inputtype=0", "iscorrection=1", "isfuzzy=0", "keyword={}".format(keyword), "mid={}".format(mid),"page=1", "pagesize=30", "platform=WebFilter", "privilege_filter=0",
                "srcappid=2919", "token=", "userid=0", "uuid={}".format(mid), "NVPh5oo715z5DIWAeQlhMDsWXXQV4hwt"]
        # 创建md5对象
        new_md5 = md5()
        info2 = ''.join(info)
        # 更新哈希对象
        new_md5.update(info2.encode(encoding='utf-8'))
        # 加密
        signature = new_md5.hexdigest()
        url = 'https://complexsearch.kugou.com/v2/search/song?appid=1014&bitrate=0&callback=callback123&clienttime={0}&clientver=1000&dfid=4XSQkz1mmSGI2XV1Ud1xgR9V&filter=10&inputtype=0&iscorrection=1&isfuzzy=0&' \
              'keyword={1}&mid={2}&page=1&pagesize=30&platform=WebFilter&privilege_filter=0&srcappid=2919&token=&userid=0&uuid={3}&signature={4}'.format(clienttime, keyword, mid, mid, signature)
        return url, mid
    # '获取网址请求'
    def get_html(self, url):
        # 加一个cookie
        cookie = 'kg_mid=7be997b45a74490a830402a853747014; Hm_lvt_aedee6983d4cfc62f509129360d6bb3d=1686481620; kg_dfid=2C8AFx0Nqndy0Quhj52Mh6EP; kg_dfid_collect=d41d8cd98f00b204e9800998ecf8427e; Hm_lpvt_aedee6983d4cfc62f509129360d6bb3d=1686483478'.split('; ')
        cookie_dict = {}
        for co in cookie:
            co_list = co.split('=')
            cookie_dict[co_list[0]] = co_list[1]
        try:
            # response = requests.get(url, headers=self.headers, cookies=cookie_dict)
            response = requests.get(url, headers=self.headers)
            response.raise_for_status()
            response.encoding = 'utf-8'
            return response.text
        except Exception as err:
            print(err)
            return '请求异常'

    def parse_text(self, text):
        count = 0
        hash_list = []
        print('{:*^80}'.format('搜索结果如下'))
        print('{0:{5}<5}{1:{5}<15}{2:{5}<10}{3:{5}<10}{4:{5}<20}'.format('序号', '歌名', '歌手', '时长(s)', '专辑', chr(12288)))
        print('{:-^84}'.format('-'))
        song_list = json.loads(text)['data']['lists']
        for song in song_list:
            print(66)
            print(song)
            print(song["ID"])
            singer_name = song['SingerName']
            # <em>本兮</em> 正则提取
            # 先匹配'</em>'这4中字符, 然后将其替换
            pattern = re.compile('[</em>]')
            singer_name = re.sub(pattern, '', singer_name)
            song_name = song['SongName']
            song_name = re.sub(pattern, '', song_name)
            album_name = song['AlbumName']
            ID = song['ID']
            # 时长
            duration = song['Duration']
            file_hash = song['FileHash']
            file_size = song['FileSize']

            # 音质为HQ, 高品质
            hq_file_hash = song['HQFileHash']
            hq_file_size = song['HQFileSize']

            # 音质为SQ, 超品质, 即无损, 后缀为flac
            sq_file_hash = song['SQFileHash']
            sq_file_size = song['SQFileSize']

            # MV m4a
            mv_hash = song['MvHash']
            m4a_size = song['M4aSize']

            hash_list.append([file_hash, hq_file_hash, sq_file_hash, ID])

            # print('{0:{5}<5}{1:{5}<15}{2:{5}<10}{3:{5}<10}{4:{5}<20}{5:{5}<20}'.format(count, song_name, singer_name, duration, album_name, ID, chr(12288)))
            print('{0:{5}<5}{1:{5}<15}{2:{5}<10}{3:{5}<10}{4:{5}<20}'.format(count, song_name, singer_name, duration, album_name, chr(12288)))
            count += 1
            if count == 10:
                # 为了测试方便, 这里只显示了10条数据
                break
        print('{:*^80}'.format('*'))
        return hash_list

    def save_file(self, song_text):
        filepath = r'F:\输出文件'
        if not os.path.exists(filepath):
            os.mkdir(filepath)
        text = json.loads(song_text)['data']
        print(77)
        print(text)
        audio_name = text['audio_name']
        author_name = text['author_name']
        album_name = text['album_name']
        img_url = text['img']
        lyrics = text['lyrics']
        play_url = text['play_url']
        print(88)
        print(play_url)
        response = requests.get(play_url, headers=self.headers)
        with open(os.path.join(filepath, audio_name) + '.mp3', 'wb') as f:
            f.write(response.content)
            print("下载完毕!")


if __name__ == '__main__':
    kg = KuGouMusic()
    search_info = input("请输入歌名或歌手: ")
    search_url, mid = kg.MD5Encrypt(search_info)
    print(11)
    print('获取请求网址：' + str(search_url))

    search_text = kg.get_html(search_url)
    print(2002)
    print(search_text[12:-2])

    hash_list = kg.parse_text(search_text[12:-2])
    print(33)
    print(hash_list)
    print(303)
    while True:
        input_index = eval(input("请输入要下载歌曲的序号(-1退出): "))
        if input_index == -1:
            break
        download_info = hash_list[input_index]
        print(44)
        print(download_info)
        # print(download_info["ID"])
        # song_url = 'https://wwwapi.kugou.com/yy/index.php?r=play/getdata&hash={}'.format(download_info[0])
        song_url = "https://wwwapi.kugou.com/yy/index.php?r=play/getdata&mid={0}&album_audio_id={1}".format(mid, download_info[3])
        song_text = kg.get_html(song_url)
        print(55)
        print(song_text)
        kg.save_file(song_text)
