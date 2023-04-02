import streamlit as st

def main():
    st.title('My Streamlit App')
    # ダウンロードボタンを作成する
    download_button = st.download_button(
        label='Download File',
        data='https://esg-ev.s3.ap-northeast-1.amazonaws.com/dist.zip',
        file_name='dist.zip'
    )

if __name__ == '__main__':
    main()




