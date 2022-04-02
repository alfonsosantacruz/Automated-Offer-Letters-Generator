### Hello Sign API Functions
from hellosign_sdk import HSClient

def send_sign_req(sign_client, filename, name, email, drafted_path):
    # Sends the Document through Hello Sign
    sign_req = sign_client.send_signature_request(
        test_mode = False,
        title = filename + ' WS AY 21-22 Offer Letter',
        subject = filename + ' WS AY 21-22 Offer Letter',
        message = 'Hi, Please review and sign the offer letter through Hello Sign as part of your Work Study position at Minerva. Best!',
        signers = [
            {'email_address': email, 'name': name},
            # CFO Email is a must in each letter. Inputs as fixed string
            # {'email_address': 'cfo@minerva.edu', 'name': 'CFOs Name'}
        ],
#         cc_email_addresses=[
#             'asantacruz@uni.minerva.edu', 
#         ],
        use_text_tags = True,
        hide_text_tags = True,
        files = [drafted_path + filename + '.pdf'])

    print('Used HelloSign to send Letter to ', name, email)
    
    return sign_req


def download_completed_offer(sign_client, sign_req, sign_req_id, completed_path):
    sign_client.get_signature_request_file(signature_request_id = sign_req_id,
                                                        filename = completed_path + sign_req.title + '.pdf',
                                                        file_type ='pdf')

    
    
def send_reminder(sign_client, sign_req_id, hire_email):
    sign_client.remind_signature_request(signature_request_id = sign_req_id, email_address = hire_email)
