"""Entry point for SD-WAN report generation."""

from catalystwan.session import create_manager_session
from mte_sdwanapi.api import (
    get_attached_device_values,
    get_device_template_definitions,
    get_feature_template_definitions,
    get_reachable_devices,
    get_vedge_policy_definitions,
    get_vmanage_credentials,
    parse_unique_device_template_ids,
    parse_unique_template_key_pairs,
    tprint,
)
from mte_sdwanapi.cli import parse_args
from mte_sdwanapi.report import save_report_to_excel


if __name__ == "__main__":

    args = parse_args()
    print_active = args.verbose
    url, username, password, port = get_vmanage_credentials()
    
    with create_manager_session(
        url=url,
        username=username,
        password=password,
        port=int(port),
    ) as session:

        if session:
            tprint("SD-WAN Manager session created successfully!")

        reachable_devices = get_reachable_devices(session, verbose=print_active)
        unique_template_ids = parse_unique_device_template_ids(reachable_devices, verbose=print_active)
        device_template_definitions = get_device_template_definitions(session, unique_template_ids, verbose=print_active)
        unique_template_key_pairs = parse_unique_template_key_pairs(device_template_definitions, verbose=print_active)
        feature_template_definitions = get_feature_template_definitions(
            session,
            unique_template_key_pairs,
            verbose=print_active,
        )
        policy_definitions = get_vedge_policy_definitions(session, verbose=print_active)
        attached_device_values = get_attached_device_values(
            session,
            device_template_definitions,
            reachable_devices,
            verbose=print_active,
        )
        save_report_to_excel(
            "device_report.xlsx",
            reachable_devices,
            device_template_definitions,
            feature_template_definitions,
            policy_definitions,
            attached_device_values,
        )

    tprint("SD-WAN Manager session closed!")

