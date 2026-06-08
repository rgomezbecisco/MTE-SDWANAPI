"""API helpers and data parsing utilities for SD-WAN reporting."""

import json
from datetime import datetime
from os import getenv

import urllib3
from pygments import formatters, highlight, lexers

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def tprint(message):
    """Print a timestamped console message."""
    timestamp = datetime.now().strftime("[%H:%M:%S %m%d%Y]")
    print(f"{timestamp} {message}")


def pretty_print_dict_as_json(data):
    """Print JSON data with indentation and terminal colors."""

    formatted_json = json.dumps(data, indent=4)
    colorful_json = highlight(
        formatted_json, lexers.JsonLexer(), formatters.TerminalFormatter()
    )

    print(colorful_json)


def get_vmanage_credentials():
    """Return vManage credentials from environment variables."""

    host = getenv("VMANAGE_IP")
    user = getenv("VMANAGE_USER")
    password = getenv("VMANAGE_PASSWORD")
    port = getenv("VMANAGE_PORT", "443")

    return host, user, password, port


def data_from_get_request(session, endpoint):
    """Return normalized JSON data from a GET request."""

    tprint("Fetching data from {}".format(endpoint))

    response = session.get(endpoint)
    payload = response.json()

    if isinstance(payload, dict) and "data" in payload:
        return payload["data"]

    if isinstance(payload, dict) and payload.get("error"):
        raise RuntimeError(
            f"API returned an error for {endpoint}: {payload.get('error')}"
        )

    return payload


def data_from_post_request(session, endpoint, payload):
    """Return normalized JSON data from a POST request."""

    tprint("Posting data to {}".format(endpoint))

    response = session.post(endpoint, json=payload)
    response_payload = response.json()

    if isinstance(response_payload, dict) and "data" in response_payload:
        return response_payload["data"]

    if isinstance(response_payload, dict) and response_payload.get("error"):
        raise RuntimeError(
            f"API returned an error for {endpoint}: {response_payload.get('error')}"
        )

    return response_payload


def parse_unique_device_template_ids(reachable_devices, verbose=False):
    """Return unique device template IDs from reachable devices."""

    unique_template_ids = set()
    for device in reachable_devices:
        template_id = device.get("templateId")
        if template_id:
            unique_template_ids.add(template_id)

    unique_template_ids_dict = {"unique_template_ids": list(unique_template_ids)}

    if verbose:
        pretty_print_dict_as_json(unique_template_ids_dict)

    return unique_template_ids_dict


def get_reachable_devices(session, verbose=False):
    """Return normalized reachable edge devices."""

    device_endpoint = "dataservice/system/device/vedges"
    device_list = data_from_get_request(session, device_endpoint)
    headers = [
        "host-name",
        "uuid",
        "system-ip",
        "site-id",
        "template",
        "templateId",
        "version",
        "reachability",
    ]
    reachable_devices = []
    for device in device_list:
        if device.get("reachability") == "reachable":
            reachable_devices.append({header: device.get(header) for header in headers})

    if verbose:
        pretty_print_dict_as_json(reachable_devices)

    return reachable_devices


def get_device_template_definitions(session, unique_template_ids, verbose=False):
    """Return full definitions for each unique device template ID."""

    if not unique_template_ids.get("unique_template_ids"):
        tprint("No unique template IDs found in reachable devices. Fetching all device templates to extract IDs...")
        all_device_template_endpoint = "dataservice/template/device"
        all_device_templates = data_from_get_request(session, all_device_template_endpoint)
        non_default_templates = [template for template in all_device_templates if not template.get("factoryDefault")]
        tprint(f"Extracted {len(non_default_templates)} non-default device templates to parse for unique IDs.")
        for template in non_default_templates:
            template_id = template.get("templateId")
            if template_id:
                unique_template_ids.setdefault("unique_template_ids", []).append(template_id)

        if verbose:
            pretty_print_dict_as_json(unique_template_ids)

    device_template_definitions = []
    for template_id in unique_template_ids["unique_template_ids"]:
        template_endpoint = f"dataservice/template/device/object/{template_id}"
        template_definition = data_from_get_request(session, template_endpoint)
        device_template_definitions.append(template_definition)

    if verbose:
        pretty_print_dict_as_json(device_template_definitions)

    return device_template_definitions


def parse_unique_template_key_pairs(device_template_definitions, verbose=False):
    """Return unique templateId/templateType pairs from nested template trees."""

    unique_templates = {}

    def walk_template_tree(node):
        if isinstance(node, dict):
            template_id = node.get("templateId")
            template_type = node.get("templateType")

            if template_id and template_type:
                unique_templates[template_id] = {
                    "templateId": template_id,
                    "templateType": template_type,
                }

            for value in node.values():
                walk_template_tree(value)

        elif isinstance(node, list):
            for item in node:
                walk_template_tree(item)

    walk_template_tree(device_template_definitions)

    unique_templates_list = sorted(
        unique_templates.values(),
        key=lambda item: item["templateId"],
    )

    result = {"unique_templates": unique_templates_list}

    if verbose:
        pretty_print_dict_as_json(result)

    return result


def get_feature_template_definitions(session, unique_template_key_pairs, verbose=False):
    """Return feature template definitions for collected template IDs."""

    feature_template_definitions = []
    for template in unique_template_key_pairs["unique_templates"]:
        template_id = template.get("templateId")
        template_type = template.get("templateType")

        if not template_id:
            continue

        feature_template_endpoint = f"dataservice/template/feature/object/{template_id}"

        try:
            feature_template_definition = data_from_get_request(
                session, feature_template_endpoint
            )
            feature_template_definitions.append(feature_template_definition)
        except RuntimeError as err:
            tprint(
                f"Failed to fetch feature template definition for {template_id}: {err}"
            )
            feature_template_definitions.append(
                {
                    "templateId": template_id,
                    "templateType": template_type,
                    "error": str(err),
                }
            )

    if verbose:
        pretty_print_dict_as_json(feature_template_definitions)

    return feature_template_definitions


def get_vedge_policy_definitions(session, verbose=False):
    """Return vEdge policy definitions."""

    policy_endpoint = "dataservice/template/policy/vedge"
    policy_definitions = data_from_get_request(session, policy_endpoint)

    if verbose:
        pretty_print_dict_as_json(policy_definitions)

    return policy_definitions


def get_attached_device_values(
    session, device_template_definitions, reachable_devices, verbose=False
):
    """Return device variable input values for each device template."""

    attached_device_values = []
    template_device_ids = {}
    for device in reachable_devices:
        template_id = device.get("templateId")
        device_id = device.get("uuid")
        if template_id and device_id:
            template_device_ids.setdefault(template_id, []).append(device_id)

    for device_template in device_template_definitions:
        template_id = device_template.get("templateId")
        if not template_id:
            continue

        device_ids = template_device_ids.get(template_id, [])
        if not device_ids:
            attached_values = []
        else:
            attached_payload = {
                "templateId": template_id,
                "deviceIds": device_ids,
                "isEdited": False,
                "isMasterEdited": False,
            }
            attached_values = data_from_post_request(
                session,
                "dataservice/template/device/config/input/",
                attached_payload,
            )

        attached_device_values.append(
            {
                "templateId": template_id,
                "attachedValues": (
                    attached_values if isinstance(attached_values, list) else []
                ),
            }
        )

    if verbose:
        pretty_print_dict_as_json(attached_device_values)

    return attached_device_values


def parse_feature_templates_for_device_template(device_template_definition):
    """Return feature template references from a device template tree."""

    feature_templates = {}

    def walk_templates(template_node):
        if not isinstance(template_node, dict):
            return

        template_id = template_node.get("templateId")
        template_type = template_node.get("templateType")
        if template_id and template_type:
            feature_templates[template_id] = {
                "templateId": template_id,
                "templateType": template_type,
            }

        for sub_template in template_node.get("subTemplates", []):
            walk_templates(sub_template)

    for template in device_template_definition.get("generalTemplates", []):
        walk_templates(template)

    return sorted(feature_templates.values(), key=lambda item: item["templateId"])

def find_and_print_default_feature_templates(feature_template_definitions, verbose=False):

    default_feature_templates = []
    for ft in feature_template_definitions:
        if ft.get("factoryDefault"):
            name = ft.get("templateName", "Unnamed Template")
            id = ft.get("templateId", "Unknown ID")
            description = ft.get("templateDescription", "No description")
            type = ft.get("templateType", "Unknown type")
            default_feature_templates.append({
                "name": name,
                "id": id,
                "description": description,
                "type": type
            })
    if default_feature_templates:
        if verbose:
            print("\nFactory Default Feature Templates:")
            pretty_print_dict_as_json(default_feature_templates)
    else:
        if verbose:
            print("\nNo factory default feature templates found.")
