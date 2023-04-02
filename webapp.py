import streamlit as st

def main():
    st.title('My Streamlit App')
    # ダウンロードボタンを作成する
    download_button = st.download_button(
        label='Download File',
        data='dist.zip',
        file_name='dist.zip'
    )

if __name__ == '__main__':
    main()




