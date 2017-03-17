"""
    This module contains function that will be used to parse information from xconfiguration
    outputs of a VCS. It does not parse everything but just those details that are not defined on the
    configuration.xml file that is within a VCS backup.

"""
import re
from pprint import pprint

# Regex to capture user administrators
regex_user_admin = r'.*Admin\sAccount\s(.*)\s(AccessAPI|AccessWeb|Enabled):\s(.*)'
# Regex to capture group administrators
regex_group_admin = r'.*Admin\sGroup\s(.*)\s(AccessAPI|AccessWeb|Enabled):\s(.*)'
# Regex to capture service configuration on the System Administration page
regex_sys_admin_services = r'.*(Administration)\s(.*)\sMode:\s(.*)'
regex_authentication_certificate = r'.*(Authentication)\s(Certificate)\sMode:\s(.*)'
# Regex to capture the Discovery Service Protection
regex_discovery_protection = r'.*(DiscoveryProtection)\s(Mode):\s(.*)'
# Regex to capture per Default DNS information
regex_perdomain_dns = r'.*(DNS\sPerDomainServer .*)\s(Address|Domain\d):\s"?([^"]*)"?'
# Regex to capture Per-domain DNS informaiton
regex_default_dns = r'.*(DNS\sServer .*)\s(Address|Index):\s"?([^"]*)"?'
# Regex to capture SIP advance information
regex_sip_advanced = r'.*(SIP\sAdvanced)\s(SdpMaxSize|SipTcpConnectTimeout):\s(.*)'
# Regex to capture the Traversal Media Port Range
regex_traversal_media_range = r'.*(Traversal\ Media).*(End|Start):\s(.*)'
# Regex to capture the SIP domains
regex_sip_domain = r'.*SIP\sDomain\s(\d*)\s(.*):\s"?([^"]*)"?'
# Regex to capture Password Security Information
regex_password_security = r'.*Authentication\s(StrictPassword)\s(.*):\s(.*)'
# Regex to capture Web interface information including session timeout and HSTS Mode configuration
regex_session_hsts = r'.*(Management)\s(Session.*|Interface.*):\s(.*)'
# Regex to capture login source information
regex_login_source = r'.*(Login\sSource)\s(.*):\s"?(\w*)"?'
#Regex to capture remote login information
regex_remote_login = r'.*(Login\sRemote)\s(Protocol|LDAP.*):\s"?(\w*)"?'
#Regex to capture DNS and IP information
regex_dns_ip = r'.*IP\s(DNS|Ephemeral|QoS)\s(.*):\s"?([^"]*)"?'
#Regex to capture QoS info from x8.9
regex_qos = r'.*\s(QoS)\s(.*):\s"?([^"|\n]*)"?'
#Regex to capture System Unit information
regex_system_unit = r'.*\s(SystemUnit)\s(.*):\s"?([^"]*)"?'
#Regex to capture IP Protocol Information
regex_ip_protocol = r'.*(xConfiguration)\s(IPProtocol):\s"?([^"]*)"?'
# Regex to capture IP Gateway Information
regex_ip_gateway = r'.*(IP)\s(Gateway|External Interface):\s"?([^"]*)"?'
# Regex to capture interface 1 configuration
regex_interface_ip_1 = r'.*(Ethernet 1)\sIP\sV\d\s(SubnetMask|Address):\s"?(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3})"?'
regex_interface_nat_1 = r'.*(Ethernet 1)\sIP\sV\d\s(StaticNAT Address|StaticNAT Mode):\s"?([^"]*)"?'
# Regex to capture interface 2 configuration
regex_interface_ip_2 = r'.*(Ethernet 2)\sIP\sV\d\s(SubnetMask|Address):\s"?(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3})"?'
regex_interface_nat_2 = r'.*(Ethernet 2)\sIP\sV\d\s(StaticNAT Address|StaticNAT Mode):\s"?([^"]*)"?'
# Regex to capture interface speed configuration
regex_interface_speed = r'.*(Ethernet)\s(1|2)\sSpeed:\s"?(\w*)"?'
# Regex to capture Static Routes
regex_static_routes = r'.*(IP Route\s\d*)\s(.*):\s"?(.*)"?'
# Regex to capture NTP information
regex_ntp_server = r'.*(NTP Server|TimeZone)\s(.*):\s"?([^"]*)"?'
# Regex to capture SNMP information
regex_snmp = r'.*(SNMP)\s(.*):\s"?([^"\n]*)"?'
# Regex to capture Clustering info
regex_clustering = r'.*(Alternates)\s(Peer \d|Cluster Name|ConfigurationMaster).*:\s"?([^"\n]*)"?'
# Regex to capture ExternalManager info
regex_ext_manager = r'.*(ExternalManager)\s(.*):\s"?([^"\n]*)"?'
# Regex to capture H323 information
regex_h323_conf = r'.*(Gatekeeper|Gateway)\s(.*):\s"?([^"\n]*)"?'
regex_h323_mode = r'.*(H323)\s(Mode):\s"?([^"\n]*)"?'
# Regex to capture SIP information
regex_sip_mode = r'.*(SIP)\s(Mode):\s"?([^"\n]*)"?'
regex_sip_info = r'.*SIP\s(Authentication|MTLS|Require|Session|TCP|UDP|TLS|MediaRouting|Registration)\s(.*):\s"?([^"\n]*)"?'
# Regex to capture interworking mode
regex_interworking = r'.*(Interworking)\s(Mode):\s"?([^"\n]*)"?'
# Regex to capture registration information
regex_registration_policy = r'.*\sRegistration\s(RestrictionPolicy)\s(Mode):\s"?([^"\n]*)"?'
regex_registration_allow_deny = r'.*(AllowList\s\d*|DenyList\s\d*)\s(.*):\s"?([^"\n]*)"?'
regex_registration_policy_service = r'.*RestrictionPolicy\s(Service\s\d*)\s(.*):\s"?([^"\n]*)"?'
# Authentication information
regex_outbound_credentials = r'.*(Authentication)\s(UserName):\s(.*)'
regex_ads = r'.*(ADS)\s(.*):\s"?([^"\n]*)"?'
regex_ntlm = r'.*(NTLM)\s(Mode):\s"?([^"\n]*)"?'
regex_h350 = r'.*(H350|LDAP)\s(Bind.*|Ldap.*|Mode|Directory.*|AliasOrigin):\s"?([^"\n]*)"?'
#Regex to capture call routing info
regex_call = r'.*Call\s(Loop|Routed|Services)\s(.*):\s"?([^"\n]*)"?'
# Regex DefaultSubzone, TraversalSubzone
regex_default_subzone = r'.*\s(DefaultSubZone)\s(.*):\s"?([^"\n]*)"?'
regex_traversal_subzone = r'.*\s(TraversalSubZone)\s(.*):\s"?([^"\n]*)"?'
# Regex Subzone and Membership
regex_subzones = r'.*\sSubZones\s(SubZone\s\d*)\s(.*):\s"?([^"\n]*)"?'
regex_membership = r'.*\sSubZones\s(MembershipRules Rule\s\d*)\s(.*):\s"?([^"\n]*)"?'
# Regex Default Zone
regex_def_zone = r'.*\s(DefaultZone)\s(.*):\s"?([^"\n]*)"?'
# Regex to capture zone information
regex_zone_type =  r'.*Zones\s(Zone\s\d+)\s(Type):\s"?([^"\n]*)"?'
regex_zone_config = r'.*Zones\s(Zone\s\d*)\s(.*):\s"?([^"\n]*)"?'
# Regex to capture transforms and Search Rules
regex_transform = r'.*\s(Transform\s\d*)\s(.*):\s"?([^"\n]*)"?'
regex_search = r'.*\sPolicy\sSearchRules\s(Rule\s\d*)\s([\s\w]*):\s"?([^"\n]*)"?'
# Regex Policy services
regex_policy_services = r'.*\sPolicy\sServices\s(Service\s\d*)\s([\s\w\d]*):\s"?([^"\n]*)"?'
# Regex for CAC
regex_bw = r'.*\s(Bandwidth)\s(D.*):\s"?([^"\n]*)"?'
regex_bw_link = r'.*\sBandwidth\s(Link\s\d*)\s(.*):\s"?([^"\n]*)"?'
regex_bw_pipe = r'.*\sBandwidth\s(Pipe\s\d*)\s(.*):\s"?([^"\n]*)"?'
# Regex for Applications
regex_multiway = r'.*\sApplications\s(ConferenceFactory)\s(.*):\s"?([^"\n]*)"?'
regex_presence = r'.*\sApplications\sPresence\s(Server|User)\s(.*):\s"?([^"\n]*)"?'
regex_findme = r'.*\sPolicy\s(FindMe)\s(.*):\s"?([^"\n]*)"?'
# Regex for Maintenance Mode
regex_maintenance = r'.*SystemUnit\s(Maintenance)\s(Mode):\s"?([^"\n]*)"?'
# Regex to obtain localdatabase credentials
regex_localDB = r'.*\s(Credential\s\d*)\s(.*):\s"?([^"\n]*)"?'
# Regex Traversal info
regex_traversal_ports = r'.*\sTraversal\s(Media|Server\sH323|Server\sMedia)\s(.*):\s"?([^"\n]*)"?'
regex_turn = r'.*\sTraversal\sServer\s(TURN)\s(.*):\s"?([^"\n]*)"?'
regex_traversal_endpoints = r'.*LocalZone\sTraversal\sH323\s(.*)\s(.*):\s"?([^"\n]*)"?'
regex_h323_pref = r'.*LocalZone\sTraversal\s(H323)\s(Preference):\s"?([^"\n]*)"?'

# Regex for Collaboration Edge
regex_collab_edge= r'.*\s(CollaborationEdge)\s(.*):\s"?([^"\n]*)"?'
regex_collab_edge_deployments= r'.*\sCollaborationEdgeDeployments\s(\d*)\s(.*):\s"?([^"\n]*)"?'



def xconf_to_dict(file_name,regex):
    """
    Function that takes the information from an xconfiguration file and organize it within a dictionary
    for easy processing
    :param file_name: File that contains the xconfiguration output
    :param regex: Regex expression used to match the options. It should have three parentesis to capture the information
    :return: dictionary of dictionary {match.group(1){match.group(2):match.group(3)}}
    """

    dictionary = {}
    funct_regex = re.compile(regex)
    with open(file_name, mode='r', encoding= "utf-8") as file:
        for lines in file:
            if not lines.startswith("*c"):
                continue
            else:
                try:
                    match = funct_regex.search(lines)
                    if match:
                        try:
                            dictionary[match.group(1)][match.group(2)] = match.group(3)
                        except:
                            dictionary[match.group(1)] = {}
                            dictionary[match.group(1)][match.group(2)] = match.group(3)
                except:
                    print ("No match in line")
    # pprint (dictionary)
    return dictionary





# Run if module is run as separate code
if __name__ == '__main__':
    file = '/Users/josemerchan/Documents/Advanced_Services/Projects/2017/Statoil/Q3/VCS audit review/ST 2/xconf_vcsc08.txt'
    #pprint (xconf_to_dict(file,regex_user_admin))
    #pprint (xconf_to_dict(file,regex_group_admin))
    #pprint (xconf_to_dict(file,regex_sys_admin_services))
    #pprint (xconf_to_dict(file,regex_discovery_protection))
    #pprint (xconf_to_dict(file,regex_perdomain_dns))
    #pprint (xconf_to_dict(file,regex_default_dns))
    #pprint (xconf_to_dict(file,regex_sip_advanced))
    #pprint (xconf_to_dict(file,regex_traversal_media_range))
    #pprint (xconf_to_dict(file,regex_sip_domain))
    #pprint (xconf_to_dict(file,regex_password_security))
    #pprint (xconf_to_dict(file,regex_session_hsts))
    #pprint (xconf_to_dict(file,regex_login_source))
    #pprint (xconf_to_dict(file,regex_remote_login))
    #pprint (xconf_to_dict(file,regex_dns_ip))
    #pprint(xconf_to_dict(file,regex_system_unit))
    #pprint(xconf_to_dict(file,regex_ip_protocol))
    #pprint (xconf_to_dict(file,regex_ip_gateway))
    #pprint(xconf_to_dict(file,regex_interface_ip_1))
    #pprint(xconf_to_dict(file,regex_interface_ip_2))
    #pprint(xconf_to_dict(file,regex_interface_nat_1))
    #pprint(xconf_to_dict(file,regex_interface_nat_2))
    #pprint(xconf_to_dict(file,regex_interface_speed))
    #pprint(xconf_to_dict(file,regex_static_routes))
    #pprint (xconf_to_dict(file,regex_ntp_server))
    #pprint (xconf_to_dict(file,regex_snmp))
    #pprint (xconf_to_dict(file,regex_clustering))
    #pprint (xconf_to_dict(file,regex_ext_manager))
    #pprint (xconf_to_dict(file,regex_h323_mode))
    #pprint (xconf_to_dict(file,regex_h323_conf))
    #pprint (xconf_to_dict(file,regex_sip_mode))
    #pprint (xconf_to_dict(file,regex_sip_info))
    #pprint (xconf_to_dict(file,regex_interworking))
    #pprint (xconf_to_dict(file,regex_registration_policy))
    #pprint (xconf_to_dict(file,regex_registration_allow_deny))
    #pprint (xconf_to_dict(file,regex_registration_policy_service))
    #pprint (xconf_to_dict(file,regex_outbound_credentials))
    #pprint (xconf_to_dict(file,regex_ads))
    #pprint (xconf_to_dict(file,regex_ntlm))
    #pprint (xconf_to_dict(file,regex_h350))
    #pprint (xconf_to_dict(file,regex_call))
    #pprint(xconf_to_dict(file,regex_subzones))
    #pprint(xconf_to_dict(file,regex_snmp))
    #pprint(xconf_to_dict(file,regex_transform))
    #pprint(xconf_to_dict(file,regex_search))
    #pprint(xconf_to_dict(file,regex_policy_services))
    #pprint(xconf_to_dict(file,regex_bw))
    #pprint(xconf_to_dict(file,regex_bw_link))
    #pprint(xconf_to_dict(file,regex_bw_pipe))
    #pprint(xconf_to_dict(file,regex_multiway))
    #pprint(xconf_to_dict(file,regex_presence))
    #pprint(xconf_to_dict(file,regex_traversal_endpoints))
    #pprint(xconf_to_dict(file,regex_h323_pref))
    #pprint(xconf_to_dict(file,regex_collab_edge))
    pprint(xconf_to_dict(file,regex_zone_type))
    #pprint(xconf_to_dict(file, regex_zone_config))

