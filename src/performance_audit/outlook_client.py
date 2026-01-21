import win32com.client as win32
import pywintypes

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"


class OutlookClient:
    def __init__(self, from_smtp: str, store_hint: str = "riscos"):
        self.from_smtp = from_smtp
        self.store_hint = store_hint

        self.outlook = win32.Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.ns.Logon("", "", False, False)

        self.account = self._find_account(from_smtp)
        if not self.account:
            raise RuntimeError(f"Conta '{from_smtp}' não encontrada no Outlook clássico.")

    def _find_account(self, smtp: str):
        for acc in self.outlook.Session.Accounts:
            try:
                if str(acc.SmtpAddress).lower() == smtp.lower():
                    return acc
            except Exception:
                continue
        return None

    def get_sent_folder(self):
        try:
            for i in range(1, self.ns.Stores.Count + 1):
                store = self.ns.Stores.Item(i)
                try:
                    if store and store.DisplayName and self.store_hint.lower() in store.DisplayName.lower():
                        return store.GetDefaultFolder(5)  # 5 = Sent Items
                except Exception:
                    continue
        except pywintypes.com_error:
            pass

        return self.ns.GetDefaultFolder(5)

    def create_mail(self):
        mail = self.outlook.CreateItem(0)
        try:
            mail.SendUsingAccount = self.account
        except Exception:
            pass
        try:
            mail.SentOnBehalfOfName = self.from_smtp
        except Exception:
            pass
        return mail

    def extract_ids(self, mail_item) -> dict:
        entry_id = getattr(mail_item, "EntryID", "") or ""
        conversation_id = getattr(mail_item, "ConversationID", "") or ""
        internet_msg_id = ""
        try:
            pa = mail_item.PropertyAccessor
            internet_msg_id = pa.GetProperty(PR_INTERNET_MESSAGE_ID) or ""
        except Exception:
            internet_msg_id = ""
        return {
            "entry_id": entry_id,
            "conversation_id": conversation_id,
            "internet_message_id": internet_msg_id
        }
