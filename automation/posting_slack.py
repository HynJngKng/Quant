

from slacker import Slacker




def slack_alert(date, work_name, elapsed_time, channel, post_script=None):
    slackClient = Slacker("xoxb-427665729104-437256985204-ZH2WXkeI3gvJgp2zSJWVj6yW") # TRST_QUANT TOKEN

    messageToChannel = str(date) + " " + str(work_name) + " Updated! elapsed " + str(elapsed_time) + "m"
    messageToChannel = messageToChannel + " " + str(post_script)

    slackClient.chat.post_message(channel, messageToChannel, as_user=True)


if __name__ == '__main__':

    date = 20180920
    work_name = "test_SlackAlert"
    elapsed_time = 2
    channel = '#daily'
    post_script = "for test"


    slack_alert(date, work_name, elapsed_time, channel, post_script=post_script)

    # 'date' 'work_name' Updated! elapsed 'elapsed_time'm
    # 위와같은 포맷으로 channel에 posting 된다
    # post script를 쓰면 저 문구 뒤에 추가해서 스크립트가 나온다