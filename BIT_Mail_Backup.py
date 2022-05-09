import time
from Mail_Backup_Core import backup_config, save_mailbox


def main():
    User, Passwd, Email_box, Batch_size, Imap_url, Port = backup_config()

    print("Wait for the mailbox server to respond...")

    start_time = time.time()

    for email_box in Email_box.split(','):
        # print(email_box, type(email_box))
        save_mailbox(Imap_url, Port, User, Passwd, email_box, Batch_size)

    end_time = time.time()
    print(f"共耗时{int((end_time - start_time) // 60)}分{int((end_time - start_time) % 60)}秒")
    input("Close the window or press Enter to quit...")


if __name__ == "__main__":
    main()
