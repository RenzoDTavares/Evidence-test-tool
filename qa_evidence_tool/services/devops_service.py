import base64
import requests
import xml.etree.ElementTree as ET
import re
import html
from pathlib import Path
from core.config import logger, CONFIG

class AzureDevOpsService:
    def __init__(self, personal_access_token):
        self.pat = personal_access_token
        self.auth_headers = {
            'Content-Type': 'application/json-patch+json',
            'Authorization': f'Basic {base64.b64encode(f":{self.pat}".encode()).decode()}'
        }
        self.base_url = CONFIG["ORGANIZATION_URL"]
        self.timeout = CONFIG["API_TIMEOUT"]

    def get_test_steps(self, test_case_id):
        url = f"{self.base_url}/_apis/wit/workitems/{test_case_id}?$select=Microsoft.VSTS.TCM.Steps&api-version=6.0"
        try:
            logger.info(f"Fetching steps for TC {test_case_id}")
            response = requests.get(url, headers={'Authorization': self.auth_headers['Authorization']}, timeout=self.timeout)
            response.raise_for_status()
            work_item = response.json()
            steps_xml = work_item['fields'].get('Microsoft.VSTS.TCM.Steps', '<steps></steps>')
            return self._parse_devops_steps(steps_xml)
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to fetch steps: {e}")
            raise ConnectionError(f"Failed to fetch steps for TC {test_case_id}: {str(e)}")

    def _parse_devops_steps(self, steps_xml):
        parsed_data = {'pre_requisite': [], 'steps': []}
        try:
            root = ET.fromstring(steps_xml)
        except ET.ParseError:
            logger.warning("XML Parse Error in DevOps steps.")
            return parsed_data

        def clean_text_content(element):
            if element is None: return ""
            try:
                raw_content = ET.tostring(element, encoding='unicode')
            except TypeError:
                return ""
            content = re.sub(r'^<[^>]+>', '', raw_content) 
            content = re.sub(r'<\/[^>]+>$', '', content)  
            content = html.unescape(content)
            content = re.sub(r'<[^>]+>', ' ', content)
            return ' '.join(content.split()).strip()

        for step_element in root.findall('.//step'):
            params = step_element.findall('parameterizedString')
            action_element = params[0] if len(params) > 0 else None
            expected_element = params[1] if len(params) > 1 else None
            
            step_text = clean_text_content(action_element)
            expected_text = clean_text_content(expected_element)

            lower_step = step_text.lower()
            if 'pre-requisite' in lower_step or 'pré-requisito' in lower_step or 'prereq' in lower_step:
                parsed_data['pre_requisite'].append(step_text)
                continue 

            if not step_text: step_text = "Action (No description)"
            if not expected_text: expected_text = "Expected Result (No description)"

            parsed_data['steps'].append({'step': step_text, 'expected': expected_text})
            
        if parsed_data['pre_requisite']:
            parsed_data['pre_requisite'] = " | ".join(parsed_data['pre_requisite'])
        else:
            parsed_data['pre_requisite'] = None
            
        return parsed_data

    def get_projects(self):
        url = f"{self.base_url}/_apis/projects?api-version=6.1-preview.4"
        try:
            logger.info("Fetching DevOps projects list.")
            response = requests.get(url, headers={'Authorization': self.auth_headers['Authorization']}, timeout=self.timeout)
            
            # Checagem defensiva: O DevOps retornou login (HTML) em vez de JSON?
            if 'json' not in response.headers.get('Content-Type', ''):
                logger.error(f"Unexpected response content: {response.text[:200]}")
                raise ValueError("DevOps returned an HTML page instead of data. Check if your ORGANIZATION_URL is correct and your PAT is valid.")
                
            response.raise_for_status()
            return [project['name'] for project in response.json().get("value", [])]
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to fetch projects: {e}")
            raise ConnectionError(f"Failed to fetch projects: {str(e)}")
        except ValueError as ve:
            raise ConnectionError(str(ve))
        
    def get_work_items_info(self, project_name, work_item_ids):
        """Fetches Titles for a list of Bug IDs individually to prevent batch failure."""
        if not work_item_ids: return {}
        
        results = {}
        
        for item_id in work_item_ids:
            # Buscando apenas o ID e o Título agora
            url = f"{self.base_url}/{project_name}/_apis/wit/workitems/{item_id}?$select=System.Id,System.Title&api-version=6.0"
            
            try:
                logger.info(f"Fetching title for bug: {item_id} in project {project_name}")
                response = requests.get(url, headers={'Authorization': self.auth_headers['Authorization']}, timeout=self.timeout)
                response.raise_for_status()
                
                item = response.json()
                title = item['fields'].get('System.Title', 'Title Not Found')
                
                # Salvando APENAS o título (removemos o prefixo "Bug: ")
                results[str(item_id)] = title
                
            except requests.exceptions.RequestException as e:
                logger.warning(f"Failed to fetch bug details for ID {item_id}: {e}")
                continue 
                
        return results

    def get_test_case_relations(self, test_case_id):
        url = f"{self.base_url}/_apis/wit/workitems/{test_case_id}?$expand=relations&api-version=6.0"
        try:
            response = requests.get(url, headers={'Authorization': self.auth_headers['Authorization']}, timeout=self.timeout)
            response.raise_for_status()
            work_item = response.json()
            test_case_title = work_item['fields'].get('System.Title', test_case_id)
            us_title = ""

            for rel in work_item.get('relations', []):
                if rel['rel'] == 'Microsoft.VSTS.Common.TestedBy-Reverse':
                    us_url = rel['url']
                    us_response = requests.get(us_url, headers={'Authorization': self.auth_headers['Authorization']}, timeout=self.timeout)
                    us_response.raise_for_status()
                    us_item = us_response.json()
                    if us_item['fields']['System.WorkItemType'] in ['Product Backlog Item', 'User Story']:
                        us_title = us_item['fields']['System.Title']
                        break
            return us_title, test_case_title
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to fetch relations for TC {test_case_id}: {e}")
            raise ConnectionError(f"Failed to fetch relations for TC {test_case_id}: {str(e)}")

    def upload_attachment(self, file_path, project_name):
        file_name = Path(file_path).name
        url = f"{self.base_url}/{project_name}/_apis/wit/attachments?fileName={file_name}&api-version=6.0"
        try:
            logger.info(f"Uploading attachment: {file_name}")
            with open(file_path, 'rb') as f:
                file_content = f.read()
            headers = {
                'Content-Type': 'application/octet-stream',
                'Authorization': self.auth_headers['Authorization']
            }
            response = requests.post(url, headers=headers, data=file_content, timeout=self.timeout * 3)
            response.raise_for_status()
            return response.json().get('url')
        except requests.exceptions.RequestException as e:
            logger.error(f"Upload failed: {e}")
            raise ConnectionError(f"Failed to upload attachment: {str(e)}")

    def link_attachment_to_work_item(self, work_item_id, attachment_url, project_name):
        url = f"{self.base_url}/{project_name}/_apis/wit/workitems/{work_item_id}?api-version=6.0"
        patch_document = [{
            "op": "add",
            "path": "/relations/-",
            "value": {
                "rel": "AttachedFile",
                "url": attachment_url,
                "attributes": {"comment": "Test evidence generated by Automation Assistant."}
            }
        }]
        try:
            logger.info(f"Linking attachment to TC {work_item_id}")
            response = requests.patch(url, headers=self.auth_headers, json=patch_document, timeout=self.timeout)
            response.raise_for_status()
            return True
        except requests.exceptions.RequestException as e:
            logger.error(f"Failed to link attachment: {e}")
            raise ConnectionError(f"Failed to link attachment: {str(e)}")