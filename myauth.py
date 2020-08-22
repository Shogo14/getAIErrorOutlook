from O365 import Account

import constants


credentials = (constants.CLIENT_ID, constants.CLIENT_SECRET)
account = Account(credentials)
account.authenticate(scopes=['basic', 'message_all'])