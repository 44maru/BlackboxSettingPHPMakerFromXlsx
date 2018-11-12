import logging.config
import os
import sys

import jctconv

LOG_CONF = "./logging.conf"
logging.config.fileConfig(LOG_CONF)

from kivy.app import App
from kivy.config import Config

Config.set('modules', 'inspector', '')  # Inspectorを有効にする
Config.set('graphics', 'width', 480)
Config.set('graphics', 'height', 280)
Config.set('graphics', 'maxfps', 20)  # フレームレートを最大で20にする
Config.set('graphics', 'resizable', 0)  # Windowの大きさを変えられなくする
Config.set('input', 'mouse', 'mouse,disable_multitouch')
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.core.window import Window
from kivy.resources import resource_add_path
from kivy.uix.screenmanager import Screen

from xlrd import open_workbook
from xlrd import XL_CELL_TEXT

if hasattr(sys, "_MEIPASS"):
    resource_add_path(sys._MEIPASS)

EMPTY = ""
SIZE_S = "S"
SIZE_M = "M"
SIZE_L = "L"
SIZE_XL = "XL"
INDEX_TWITTER = 0
INDEX_ITEM_NO = 1
INDEX_ITEM_SIZE = 2
INDEX_LAST_NAME = 3
INDEX_FIRST_NAME = 4
INDEX_POST_CODE = 5
INDEX_STATE = 6
INDEX_CITY = 7
INDEX_ADDRESS = 8
INDEX_PHONE = 9
INDEX_EMAIL = 10
INDEX_PAY_TYPE = 11
INDEX_CARD_NUMBER = 12
INDEX_CARD_LIMIT_MONTH = 13
INDEX_CARD_LIMIT_YEAR = 14
INDEX_CARD_CVV = 15
INDEX_DELAY = 16
INDEX_START_TIME = 17
INDEX_CACHE = 18

ID_MESSAGE = "message"

OUT_FILE_NAME = "blackbox-setting.php"
CHECKOUT_PROFILES_JSON = "checkoutprofiles.json"
UTF8 = "utf8"
SJIS = "sjis"

CONFIG_TXT = "./config.txt"
PROXY_TXT = "./proxy.txt"
CONFIG_DICT = {}
CONFIG_KEY_DELAY = "DELAY"
CONFIG_KEY_START_WEEK = "START_WEEK"
CONFIG_KEY_START_HHMM = "START_HHMM"
CONFIG_KEY_SECRET = "SECRET_KEY"
CONFIG_KEY_DISCORD_HOOK_URL = "webhookURL"
CONFIG_KEY_DISCORD_MESSAGE = "discordmessage"
PROXY_LIST = []

OUT_FILE_CONTENTS_HEADER = """<?php
// 
// blackbox-setting.php
// http://buzz-coin.work/blackbox-setting-v3.php?id=1 指定された設定データを $secretで暗号化した文字列で返す
// http://buzz-coin.work/blackbox-setting-v3.php?id=2
//
// http://buzz-coin.work/blackbox-setting-v2.php?id=2&mode=check の場合はproxyが指定通りかチェック 
//

mb_language("Japanese");
mb_internal_encoding('UTF-8');
header('Content-Type: text/html; charset=utf-8');

$settings = array();
"""

OUT_FILE_CONTENTS_TEMPLATE = """
$setting = array();
$setting["secret"]		= "{}";
$setting["codes1"]		= "{}";
$setting["sizes1"]		= "{}";
$setting["codes2"]		= "{}";
$setting["sizes2"]		= "{}";
$setting["codes3"]		= "{}";
$setting["sizes3"]		= "{}";
$setting["proxy"]		= "{}";
$setting["start_week"]	= {};
$setting["start_hhmm"]	= "{}";
$setting["last_name"]	= "{}";
$setting["first_name"]	= "{}";
$setting["email"]		= "{}";
$setting["tel"]			= "{}";
$setting["pref"]		= " {}";
$setting["address"]		= "{}";
$setting["address2"]	= "{}";
$setting["zip"]			= "{}";
$setting["card_type"] 	= "{}";
$setting["card_number"]	= "{}";
$setting["card_month"]	= "{}";
$setting["card_year"]	= "{}";
$setting["vval"]		= "{}";
$setting["cash"]		= {};
$setting["delay"]		= {};
$setting["discordhookurl"] = "{}";
$setting["discordmessage"] = "{}";
$settings[{}] = $setting;
"""

OUT_FILE_CONTENTS_HOOTER = """///////////////////////////////////
//
// 1.パラメーターチェック
//
///////////////////////////////////
extract($_GET);
if(!compact('id')){
	echo "error #1"; exit;
}
if(!array_key_exists($id,$settings)){
	echo "error #2 no setting id={$id}"; exit;
}

///////////////////////////////////
//
// 2. checkモード処理
//
///////////////////////////////////
if(compact('mode') && $mode=="check"){

	$proxy = explode(",",$settings[$id]["proxy"]);
	if(count($proxy)<=1)$proxy = explode(":",$settings[$id]["proxy"]);

	if($_SERVER["HTTP_X_REAL_IP"]==$proxy[0]){
		echo "OK!! your ip is {$_SERVER["HTTP_X_REAL_IP"]}";
	}elseif($_SERVER["REMOTE_ADDR"]==$proxy[0]){
		echo "OK!! your ip is {$_SERVER["REMOTE_ADDR"]}";
	}else{
		echo "Not Good! Your ip {$_SERVER["REMOTE_ADDR"]} is defferent from setting id={$id}";
	}

}else{
///////////////////////////////////
//
// 3. 設定データ返送処理
//
///////////////////////////////////

	$out = array();
	foreach($settings[$id] as $key => $value){
		$out[] = "{$key}={$value}";
	}
	$data_plain = json_encode($out);
	//$data_plain = json_encode($settings[$id]);
	//echo "<html><head></head><body><data>". $data_plain ."</data></body></html>";
	$encrypted = CryptoJSAesEncrypt($settings[$id]["secret"],$data_plain);
	echo "<html><head></head><body><data>". $encrypted ."</data></body></html>";
}

exit;


function CryptoJSAesEncrypt($passphrase, $plain_text){

    $salt	= openssl_random_pseudo_bytes(256);
    $iv		= openssl_random_pseudo_bytes(16);

    $iterations = 999;  
    $key = hash_pbkdf2("sha512", $passphrase, $salt, $iterations, 64);

    $encrypted_data = openssl_encrypt($plain_text, 'aes-256-cbc', hex2bin($key), OPENSSL_RAW_DATA, $iv);

    $data = array("ciphertext" => base64_encode($encrypted_data), "iv" => bin2hex($iv), "salt" => bin2hex($salt));
    return json_encode($data);
}

?>"""


class JsonMakerScreen(Screen):
    def __init__(self, **kwargs):
        super(JsonMakerScreen, self).__init__(**kwargs)
        self._file = Window.bind(on_dropfile=self._on_file_drop)
        self.proc_line_number = 0

    def _on_file_drop(self, window, file_path):
        self.dump_out_file(file_path.decode(UTF8))

    def dump_out_file(self, file_path):
        try:
            self.dump_out_file_core(file_path)
        except Exception as e:
            self.disp_messg_err("{}の出力に失敗しました。処理Excel行番号={}。".format(
                OUT_FILE_NAME, self.proc_line_number))
            log.exception("{}の出力に失敗しました。処理Excel行番号={}。%s".format(
                OUT_FILE_NAME, self.proc_line_number), e)

    @staticmethod
    def split_list(row, index):
        try:
            return row[index].value.split("&")
        except AttributeError:
            size = float(row[index].value)
            if size % 1.0 == 0.0:
                size = int(size)
            return [str(size)]

    def dump_out_file_core(self, file_path):
        index = 1
        proxy_index = 0

        with open(OUT_FILE_NAME, "w", encoding=UTF8) as f:

            f.write(OUT_FILE_CONTENTS_HEADER)

            workbook = open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                self.proc_line_number = i + 1
                row = sheet.row(i)
                log.info("{}行目 => {}".format(i + 1, row))

                if self.is_not_address_record(row):
                    log.info("{}行目に必須項目未入力のセルがありました。この行の取り込みをスキップします".format(i + 1))
                    continue

                item_no_list = self.split_list(row, INDEX_ITEM_NO)
                size_list = self.split_list(row, INDEX_ITEM_SIZE)

                if len(item_no_list) != len(size_list):
                    self.disp_messg_err("{}行目のアイテム数とサイズの数が一致しません。\n出力を中断しました。".format(i + 1))
                    log.error("{}行目のアイテム数とサイズの数が一致しません。アイテム:{} サイズ:{}".format(
                        i + 1, item_no_list, size_list))
                    return False

                if len(item_no_list) > 3:
                    self.disp_messg_err("{}行目のアイテム数とサイズの数が4つ以上指定されています。\n出力を中断しました。".format(i + 1))
                    log.error("{}行目のアイテム数とサイズの数が4つ以上指定されています。アイテム:{} サイズ:{}".format(
                        i + 1, item_no_list, size_list))
                    return False

                item_code_1 = item_no_list[0]
                item_size_1 = self.format_size(size_list[0])

                if len(item_no_list) > 1:
                    item_code_2 = item_no_list[1]
                    item_size_2 = self.format_size(size_list[1])
                else:
                    item_code_2 = EMPTY
                    item_size_2 = EMPTY

                if len(item_no_list) > 2:
                    item_code_3 = item_no_list[2]
                    item_size_3 = self.format_size(size_list[2])
                else:
                    item_code_3 = EMPTY
                    item_size_3 = EMPTY

                last_name = row[INDEX_LAST_NAME].value
                first_name = row[INDEX_FIRST_NAME].value
                email = row[INDEX_EMAIL].value
                phone_number = row[INDEX_PHONE].value
                state = row[INDEX_STATE].value
                city = row[INDEX_CITY].value
                detail_address = row[INDEX_ADDRESS].value
                zip_code = self.get_val_according_to_cell_type(row, INDEX_POST_CODE)
                card_type = row[INDEX_PAY_TYPE].value.lower().replace(" ", "_").replace("mastercard", "master")

                if card_type == "代金引換":
                    card_type = "visa"
                    card_number = "0000000000000000"
                    card_limit_month = "01"
                    card_limit_year = "2020"
                    cvv = "111"
                else:
                    card_number = row[INDEX_CARD_NUMBER].value
                    card_limit_month = "%02d" % row[INDEX_CARD_LIMIT_MONTH].value
                    card_limit_year = "20" + str(int(row[INDEX_CARD_LIMIT_YEAR].value))
                    cvv = self.get_val_according_to_cell_type(row, INDEX_CARD_CVV)

                delay = self.get_val_if_empty_as_default(row, INDEX_DELAY, CONFIG_DICT[CONFIG_KEY_DELAY])
                start_hhmm = self.get_val_if_empty_as_default(row, INDEX_START_TIME, CONFIG_DICT[CONFIG_KEY_START_HHMM])
                cache = self.get_val_if_empty_as_default(row, INDEX_CACHE, "true")

                if len(PROXY_LIST) <= proxy_index:
                    proxy_index = 0

                proxy = self.get_proxy_info(proxy_index)

                f.write(OUT_FILE_CONTENTS_TEMPLATE.format(
                    CONFIG_DICT[CONFIG_KEY_SECRET],
                    item_code_1, item_size_1, item_code_2, item_size_2, item_code_3, item_size_3,
                    proxy, CONFIG_DICT[CONFIG_KEY_START_WEEK],
                    start_hhmm, last_name, first_name, email,
                    phone_number, state, city, detail_address, zip_code, card_type, card_number,
                    card_limit_month, card_limit_year, cvv, cache, delay,
                    CONFIG_DICT[CONFIG_KEY_DISCORD_HOOK_URL], CONFIG_DICT[CONFIG_KEY_DISCORD_MESSAGE],
                    index
                ))

                index += 1
                proxy_index += 1

            f.write(OUT_FILE_CONTENTS_HOOTER)

        self.disp_messg("{}を出力しました".format(OUT_FILE_NAME))

    @staticmethod
    def get_val_if_empty_as_default(row, excel_index, default_val):
        val = row[excel_index].value
        if val == EMPTY:
            val = default_val
        return val

    @staticmethod
    def get_val_according_to_cell_type(row, index):
        if row[index].ctype == XL_CELL_TEXT:
            return row[index].value
        else:
            return str(int(row[index].value))

    @staticmethod
    def is_not_address_record(row):
        if row[INDEX_TWITTER].value == EMPTY or row[INDEX_ITEM_NO].value == EMPTY \
                or row[INDEX_LAST_NAME].value == EMPTY or row[INDEX_FIRST_NAME].value == EMPTY \
                or row[INDEX_LAST_NAME].value == EMPTY or row[INDEX_POST_CODE].value == EMPTY \
                or row[INDEX_STATE].value == EMPTY or row[INDEX_CITY].value == EMPTY \
                or row[INDEX_ADDRESS].value == EMPTY or row[INDEX_PHONE].value == EMPTY \
                or row[INDEX_EMAIL].value == EMPTY or row[INDEX_PAY_TYPE].value == EMPTY:
            return True

        if row[INDEX_PAY_TYPE].value != "代金引換":
            if row[INDEX_PAY_TYPE].value == EMPTY or row[INDEX_CARD_NUMBER].value == EMPTY \
                    or row[INDEX_CARD_LIMIT_MONTH].value == EMPTY or row[INDEX_CARD_LIMIT_YEAR].value == EMPTY \
                    or row[INDEX_CARD_CVV].value == EMPTY:
                return True

        return False

    def get_proxy_info(self, proxy_index):
        if len(PROXY_LIST) > 0:
            proxy = PROXY_LIST[proxy_index]
        else:
            proxy = EMPTY
        return proxy

    def disp_messg(self, msg):
        self.ids[ID_MESSAGE].text = msg
        self.ids[ID_MESSAGE].color = (0, 0, 0, 1)

    def disp_messg_err(self, msg):
        self.ids[ID_MESSAGE].text = "{}\n詳細はログファイルを確認してください。".format(msg)
        self.ids[ID_MESSAGE].color = (1, 0, 0, 1)

    def format_size(self, sizes):
        final_size = ""
        cnt = 0
        for size in sizes.split(","):
            if cnt != 0:
                final_size += ","
            final_size += self.format_one_size(size)
            cnt += 1
        return final_size

    @staticmethod
    def format_one_size(size):
        size = jctconv.normalize(size.upper())
        if size == SIZE_S or size == "SMALL":
            return "Small"
        elif size == SIZE_M or size == "MEDIUM":
            return "Medium"
        elif size == SIZE_L or size == "LARGE":
            return "Large"
        elif size == SIZE_XL or size == "XLARGE":
            return "XLarge"
        else:
            return size


class PhpMakerApp(App):
    def build(self):
        return JsonMakerScreen()


def setup_config():
    load_config()
    load_proxy()


def load_proxy():
    if not os.path.exists(PROXY_TXT):
        return

    for line in open(PROXY_TXT, "r", encoding=UTF8):
        PROXY_LIST.append(line.replace("\n", ""))


def load_config():
    for line in open(CONFIG_TXT, "r", encoding=SJIS):
        items = line.replace("\n", "").split("=")

        if len(items) != 2:
            continue

        CONFIG_DICT[items[0]] = items[1]

    if CONFIG_DICT.get(CONFIG_KEY_SECRET) is None or CONFIG_DICT[CONFIG_KEY_SECRET] == EMPTY:
        log.error("config.txtの{}の値が未設定です。".format(CONFIG_KEY_SECRET))
        raise KeyError


if __name__ == '__main__':
    try:
        log = logging.getLogger('my-log')
        setup_config()
        LabelBase.register(DEFAULT_FONT, "ipaexg.ttf")
        PhpMakerApp().run()
    except Exception as e:
        log.exception("エラー : %s", e)
