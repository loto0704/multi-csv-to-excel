import argparse
import os
import logging
import datetime
import sys

import pandas as pd


def get_arguments() -> argparse.Namespace:
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument('-i', '--input_data', required=True, help='input file')  # 対象ファイル・フォルダ
    arg_parser.add_argument('-e', '--file_encode', required=False, default='utf-8',
                            help='input file encode')  # 取込ファイル文字コード
    arg_parser.add_argument('-dg', '--debug_mode', action='store_true', help='debug')  # デバッグモード
    return arg_parser.parse_args()


# フォルダ作成
def crt_folder(folder_name):
    executable_dir = os.path.dirname(os.path.realpath(__file__))
    crt_folder_path = os.path.join(executable_dir, folder_name).__str__()
    if not os.path.isdir(crt_folder_path):
        os.mkdir(crt_folder_path)
    return crt_folder_path


def log_setting(debug_mode: bool):
    log_folder = crt_folder('loggings')
    log_format = '%(asctime)s,%(msecs)d | %(levelname)s | %(name)s - %(message)s'
    log_filepath = os.path.join(log_folder, '{}.log'.format(datetime.datetime.now().strftime("%Y-%m-%d"))).__str__()
    logging.FileHandler(filename=log_filepath, mode='a', encoding='utf-8', delay=False)
    logging.basicConfig(
        filename=log_filepath,
        level=logging.DEBUG if debug_mode else logging.INFO,
        format=log_format,
        encoding='utf-8',
    )
    logging.info('----------ログ開始----------')


def export_excel(csv_path, excel_path, encoding):
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode='a' if os.path.exists(excel_path) else 'w') as writer:
        df = pd.read_csv(csv_path, encoding=encoding)
        file_name = os.path.basename(csv_path)
        sheet_name = os.path.splitext(file_name)[0][:31]
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    args = get_arguments()
    log_setting(args.debug_mode)
    try:
        logging.info('対象ファイル・フォルダ:{}'.format(args.input_data))
        logging.info('取込ファイル文字コード:{}'.format(args.file_encode))
        logging.info('デバッグモード:{}'.format(args.debug_mode))
        file_encode = args.file_encode
        result_folder = crt_folder(result_folder_name)  # resultフォルダの作成
        output_path = os.path.join(result_folder, f'{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
        for file in os.listdir(args.input_data):
            if file.endswith(".csv"):
                file_path = os.path.join(args.input_data, file)
                export_excel(csv_path=file_path, excel_path=output_path, encoding=file_encode)

        logging.info('出力ファイルパス：{}'.format(output_path))
        logging.info('----------ログ終了----------')

    except Exception as e:
        logging.error('main:{}'.format(e))
        print('main:{}'.format(e))
        logging.info('----------ログ終了----------')
        sys.exit(1)


if __name__ == '__main__':
    result_folder_name = 'results'
    main()
