import base64
from threading import Thread
from pathlib import Path
from core.config import logger, CONFIG
from services.devops_service import AzureDevOpsService
from services.document_service import DocumentService

class QAController:
    @staticmethod
    def get_pat_token():
        key_file = CONFIG["KEY_FILE"]
        try:
            with open(key_file, "r") as f:
                encoded_key = f.read().strip()
            return base64.b64decode(encoded_key).decode('utf-8')
        except FileNotFoundError:
            return None
        except Exception as e:
            logger.error(f"Error decoding PAT token: {e}")
            return None

    @staticmethod
    def fetch_projects_async(callback_success, callback_error):
        def task():
            pat = QAController.get_pat_token()
            if not pat:
                callback_error("Missing Token", "Please register your DevOps PAT first.")
                return
            try:
                service = AzureDevOpsService(pat)
                projects = service.get_projects()
                callback_success(projects)
            except Exception as e:
                logger.exception("Failed to fetch projects")
                callback_error("API Error", str(e))
        
        Thread(target=task, daemon=True).start()

    @staticmethod
    def generate_evidence_async(data, callback_success, callback_error, callback_finally):
        def task():
            try:
                logger.info("Starting document generation process...")
                image_dir = Path(data['img_dir'])
                us_title, test_case_title = "", data['test_id']
                formatted_bugs = data['bugs']
                test_data_dict = {'pre_requisite': None, 'steps': []}
                service = None

                if data['is_devops']:
                    pat = QAController.get_pat_token()
                    if not pat: 
                        raise ValueError("No DevOps Token found.")
                    
                    service = AzureDevOpsService(pat)
                    test_data_dict = service.get_test_steps(data['test_id'])
                    
                    fetched_us, fetched_tc = service.get_test_case_relations(data['test_id'])
                    if fetched_tc:
                        us_title = fetched_us
                        test_case_title = fetched_tc
                    else:
                        raise ValueError(f"Could not find relations for TC {data['test_id']}.")

                    if data['bugs']:
                        raw_bug_entries = [b.strip() for b in data['bugs'].split(',') if b.strip()]
                        bug_ids_to_fetch = [int(b) for b in raw_bug_entries if b.isdigit()]
                        
                        if bug_ids_to_fetch:
                            bug_info = service.get_work_items_info(data['project'], bug_ids_to_fetch)
                            
                        formatted_bug_parts = []
                        for b in raw_bug_entries:
                            if b.isdigit() and b in bug_info:
                                formatted_bug_parts.append(f"{b} - {bug_info[b]}")
                            else:
                                formatted_bug_parts.append(b)
                        
                        formatted_bugs = ", ".join(formatted_bug_parts)

                tags = CONFIG["PLACEHOLDERS"]
                placeholder_values = {
                    tags['tester']: data['tester'],
                    tags['test_id']: f"{data['test_id']} - {test_case_title}" if test_case_title != data['test_id'] else data['test_id'],
                    tags['us']: us_title,
                    tags['env']: data['env'],
                    tags['profile']: data['profile'],
                    tags['bugs']: formatted_bugs,
                    tags['result']: '☐Passed  ☒Failed' if data['bugs'] else '☒Passed  ☐Failed'
                }

                new_file_name = f"{Path(data['template_file']).stem}_{data['test_id']}.docx"
                output_path = image_dir / new_file_name
                
                DocumentService.generate_evidence_doc(
                    template_path=data['template_file'],
                    image_dir=image_dir,
                    output_path=output_path,
                    data=placeholder_values,
                    test_data=test_data_dict,
                    is_devops=data['is_devops']
                )

                if data['is_devops'] and service:
                    attachment_url = service.upload_attachment(output_path, data['project'])
                    if attachment_url:
                        service.link_attachment_to_work_item(data['test_id'], attachment_url, data['project'])
                        callback_success("Success", f"Generated and attached to TC {data['test_id']}!\nSaved at: {output_path}")
                    else:
                        callback_success("Warning", f"Generated locally, but attachment failed.\nSaved at: {output_path}", True)
                else:
                    callback_success("Success", f"Evidence generated successfully!\nSaved at: {output_path}")

            except Exception as e:
                logger.exception("Error during evidence generation")
                callback_error("Execution Error", str(e))
            finally:
                callback_finally()
                
        Thread(target=task, daemon=True).start()