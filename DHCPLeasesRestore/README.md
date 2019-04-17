# DHCPLeasesRestore

Script restores DHCP leases from DHCP logs.

Limitations: 

- DHCP logs from the last 7 days are available by default
- Script works with multiple subnets, but requires all subnets to be 255.255.255.0/24 or smaller and distinguishable by third octet. First two octets must be the same for all subnets

1. In script, enter IP address of your DHCP server into strTargetDHCPServerIP
2. Update path to DHCP logs, if stored in non-default location, in arrLogfile
3. For each subnet to process, specify third octet and first IP address for each subnet into arrSubnetScopeMapping. Example: arrSubnetScopeMapping. Example: arrSubnetScopeMapping = Array ("1", "10.0.1.0", "2", "10.0.2.192", ...)
4. Execute "DHCPLeasesRestore.vbs". File "DHCPImport.bat" will be created with import commands.
5. Execute output file "DHCPImport.bat" to import DHCP leases detected in logs
