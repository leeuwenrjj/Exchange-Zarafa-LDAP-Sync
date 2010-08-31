#!/usr/bin/python
#
# This python script synchronizes the Exchange LDAP values to Zarafa LDAP Values 
# (these can be on seperate servers but, for the time being, the DN of the users must be the same)
#
# First Release: Robert van Leeuwen <leeuwenrjj+ex-za-mig @ gmail.com> 20-10-2010
# This script was made for a migration at a customer
# Many thanks to my employer for allowing this to be GPL-ed:
# Stone-IT: Linux for (e)business
# If you have any questions/additions to the code don't hesitate to contact me. 
# 
# GPL licence
#
# Please be carefull because we are going to write to the LDAP server
# First try a few dry-runs (default) to see if you get what you expected
# This tools SYNCHRONIZES fields. You can run this tool multiple times and you
# should not end up with double mail/delegates.
#
# As with all software: you get no warrenty's whatsoever
# So don't call me if you break you're LDAP server
# Only mail me you're successtory's :)
#
# Because this script was written for a specific customer some things are not migrated.
# e.g. zarafaSharedStoreOnly, zarafaAdmin are not migrated 
# Additions that sync these fields are welcome!!
#
# The following fields will be synced:
# msExchUserAccountControl      => zarafaAccount
# msExchHideFromAddressLists:   => zarafaHidden
# mDBOverHardQuotaLimit:		=> zarafaQuotaWarn
# mDBOverQuotaLimit:        	=> zarafaQuotaSoft
# mDBStorageQuota:      		=> zarafaQuotaHard
# proxyAddresses:			    => otherMailbox
# publicDelegates: 			    => zarafaSendAsPrivilege
# 

# Import the needed modules
import sys
try:
    import ldap 
except:
    print "You are missing some modules.. Is Python LDAP on the system?"
    sys.exit()



###################################################################
# define exchange server LDAP - MUST CHANGE, DEPENDS ON ENVIRONMENT
# If you have a large environment please remember to allow large results from LDAP, for MS AD see: http://support.microsoft.com/kb/315071
###################################################################
ex_server = 'ldap://10.0.0.1'                         # Exchange/AD LDAP Server
ex_dn = 'CN=admin,DC=test,DC=local'                   # User with write access to Exchange LDAP
ex_pw = 'password'                                    # Password of Exchange LDAP user
ex_base_dn = 'OU=users,DC=test,DC=local'              # Where to look for users (and al sub-containers)
ex_filter = '(mail=testuser@test.local)'              # Filter, handy for testing on 1 user use * for all users
###################################################################
# define zarafa server LDAP - MUST CHANGE, DEPENDS ON ENVIRONMENT
###################################################################
za_server = 'ldap://10.0.0.1'                         # Zarafa/AD LDAP Server, probably same as ex_server
za_dn = 'CN=admin,DC=test,DC=local'                   # User with write access to zarafa LDAP, probably same as ex_dn
za_pw = 'password'                                    # Password of Zarafa LDAP user, probably same as ex_pw
za_base_dn = 'OU=users,DC=test,DC=local'               # Where to look for users (and al sub-containers), probably same as ex_base_dn



###################################################################
# ldap server initialization defenition, leave this as-is
ex_ldap = ldap.initialize(ex_server)
za_ldap = ldap.initialize(za_server)
###################################################################

###################################################################
# Attribute mappings, you should probably leave this as-is
#
# These are the LDAP attributes we use on the Exchange side, these attributes will be mapped to correspondig za_attrs. ORDER OF THE ATTRIBUTES SHOULD MAP EXACTLY TO za_attrs!
ex_attrs = ['sAMAccountName','msExchUserAccountControl','msExchHideFromAddressLists','mDBOverHardQuotaLimit','mDBOverQuotaLimit','mDBStorageQuota','mDBUseDefaults']
za_attrs = ['sAMAccountName','zarafaAccount','zarafaHidden','zarafaquotahard','zarafaquotasoft','zarafaquotawarn','zarafaQuotaOverride']
# These are used for syncing delegates
ex_dgs=['dn','publicDelegates']                 # Exchange LDAP attributes for delegates, dn is needed and must be first
za_dgs_attribute = 'zarafaSendAsPrivilege'      # This is the zarafa Delegate attribute
za_dgs_filter=['zarafaSendAsPrivilege']         # We also need this as a filter, should probably be same as za_dgs_attribute
# These are used for syncing mail addresses
ex_mail_attribute = ['mail','proxyAddresses']   # Exchange LDAP primary and secondary attributes for mail
za_mail_attribute = 'otherMailbox'              # Zarafa secondary mail address attributes
za_mail_attribute_filter = ['otherMailbox']     # We also need this as a filter, should probably be same as za_mail_attribute
###################################################################


###################################################################
# Code starts here, only l33t coders should edit from here :)
###################################################################
#

#bind the Exchange LDAP server
try:
    ex_ldap.bind_s( ex_dn, ex_pw )
    print "Successfully connected to Exchange AD LDAP server."
except ldap.INVALID_CREDENTIALS:
    print "Your ADDS username or password is incorrect."
    sys.exit()
except ldap.LDAPError, e:
    if type(e.message) == dict and e.message.has_key('desc'):
        print e.message['desc']
    else:
        print e
    sys.exit()

#bind the Zarafa LDAP server
try: 
    za_ldap.bind_s( za_dn, za_pw )
    print "Successfully connected to Zarafa LDAP server."
except ldap.INVALID_CREDENTIALS:
    print "Your Zarafa LDAP server username or password is incorrect."
    sys.exit()
except ldap.LDAPError, e:
    if type(e.message) == dict and e.message.has_key('desc'):
        print e.message['desc']
    else:
        print e
    sys.exit()

print "By default we will do a dry-run, only showing the changes"
apply = raw_input("Do you want to also want to apply the changes when running? y/n (n) ")
#apply = "n"


# USER ATTRIBUTE SYNCHRONISATION 
# Fill de ad_ldap_all with the ldap search from AD 
print "Synchronizing user attributes in Exchange LDAP with Zarafa LDAP"
ex_ldap_attr=ex_ldap.search_s( ex_base_dn, ldap.SCOPE_SUBTREE, ex_filter, ex_attrs )
# For the values per object (dn) do a loop
for (dn, vals) in ex_ldap_attr:
# We are going to loop through the attribute array and use i as a counter 
    i = 0
# a is the current attribute we sync     
    for a in ex_attrs:
        try:
            ex_attrs_value = vals[ex_attrs[i]][0]
# Exchange uses different values, this is converted
            if a == 'msExchUserAccountControl':
                if ex_attrs_value == '0':
                    ex_attrs_value= '1'
                if ex_attrs_value == '2':
                    ex_attrs_value= '0'
            if a == 'msExchHideFromAddressLists':
                if ex_attrs_value == 'TRUE':
                    ex_attrs_value='1'
                if ex_attrs_value == 'FALSE':
                    ex_attrs_value='0'
            if a == 'mDBUseDefaults':
                if ex_attrs_value == 'FALSE':
                    ex_attrs_value='1'
                if ex_attrs_value == 'TRUE':
                    ex_attrs_value='0'
            if a== 'mDBOverHardQuotaLimit' or a == 'mDBOverQuotaLimit' or a =='mDBStorageQuota':
                quota = int(ex_attrs_value) / 1000
                ex_attrs_value = str(quota)
# First try to compare and change if different
            try: 
                if not za_ldap.compare_s(dn, za_attrs[i], ex_attrs_value):
                    try:
                        if apply== 'y':
                            za_ldap.modify_s(dn, [(ldap.MOD_REPLACE,  za_attrs[i], ex_attrs_value)])
                        print "INFO ATTR: Succesfull change: " + za_attrs[i] + " into " + ex_attrs_value + " for user " + dn 
                    except:
                        print "ERROR ATTR: problem with change for directory server " + za_attrs[i] + " into " + ex_attrs_value + " for user " + dn
                        sys.exit() 
# If compare fails the zarafa attribute probably does not exist, try to add
            except:
                print za_attrs[i] + ex_attrs_value
                try:
                    if apply== 'y':    
                        za_ldap.modify_s(dn, [(ldap.MOD_ADD,  za_attrs[i], ex_attrs_value)])
                    print "INFO ATTR: Succesfull ADD: " + za_attrs[i] + " into " + ex_attrs_value + " for user " + dn
                except:
                    print "ERROR ATTR: problem with writing to directory server for " + za_attrs[i] + " into " + ex_attrs_value + " for user " + dn
                    sys.exit()
# if no quota's are set we get an exception that we pass
        except:
            pass
        i+=1
print "Attribute sync complete"



# Public Delegates Sync, public delegates are multi-value fields and can't be synced with the default attribute sync
# Fill de ad_ldap_all with the ldap search from AD  
print "Synchronizing delegates in Exchange LDAP with Zarafa LDAP" 
ex_ldap_dgs=ex_ldap.search_s( ex_base_dn, ldap.SCOPE_SUBTREE, ex_filter, ex_dgs ) 
# Loop through all users
for (dn, vals) in ex_ldap_dgs: 
# Loop through all Delegates
    try:
        all_dgs=vals[ex_dgs[1]]
        dgs_changed = 0
        for dgs in all_dgs:
            dgs_dn=dgs
#            dgs_dn=ldap.filter.escape_filter_chars(dgs)
# Try to compare the values, if you have an exception it probably does not have the attribute
            try:
                if not za_ldap.compare_s(dn, za_dgs_attribute , dgs):
                    dgs_changed = 1
            except:
                dgs_changed = 1
#           
#        smtp_changed = 1
        if dgs_changed == 1:
            try:
                if apply== 'y':
                    za_ldap.modify_s(dn, [(ldap.MOD_DELETE, za_dgs_attribute, None )])
                print "INFO DGS: Succesfull clear: " + za_dgs_attribute + " for user " + dn
            except: 
                pass
            for dgs_dn in all_dgs:
                try: 
                    if apply== 'y':
                        za_ldap.modify_s( dn, [(ldap.MOD_ADD, za_dgs_attribute , dgs_dn)]) 
                    print "INFO DGS: Succesfull change: added " + dgs_dn + " delegate to " + dn
                except: 
                    print "ERROR DGS: problem with writing to directory server for " + dn + " delegate to " + dn 
                    #sys.exit() 
    except:
        pass
print "delegate sync complete"





# USER MAIL FIELD SYNC
# This one is a multi-value field so we can not use the default sync above
# We can differentiate between primary addresses and secondary in Exchange because 
# the primary addresses start with SMTP (uppercase in stead of lowercase)
print "Synchronizing user email addresses in Exchange with Zarafa"
# Get the exchange smtp values and start a loop for all users
ex_ldap_smtp=ex_ldap.search_s( ex_base_dn, ldap.SCOPE_SUBTREE, ex_filter, ex_mail_attribute )
for (dn, vals) in ex_ldap_smtp:
# We fill the Zarafa smtp value for this user
#    za_filterby_dn = ldap.filter.escape_filter_chars(dn)
    zadn_smtp = za_ldap.search_s ( dn, ldap.SCOPE_SUBTREE, 'mail=*', za_mail_attribute_filter )
# Count the number of email addresses in Zarafa
    za_smtp = str(zadn_smtp[0][1])
    za_smtp_number=0
    for char in za_smtp:
        if char == '@':
            za_smtp_number+=1
    smtp_ex_value = vals[ex_mail_attribute[1]]
# Initialize some variables
    ex_smtp_number = 0
    smtp_changed = 0
# Loop through the Exchange smtp fields, count the addresses and check if they are in Zarafa
    for ex_smtp in smtp_ex_value:
        if ex_smtp[0:4] == 'smtp':
            ex_smtp_number+=1   
            try:
                if not za_ldap.compare_s(dn, za_mail_attribute , ex_smtp[5:]):
                    smtp_changed=1 
            except:
                pass
# If a mailaddress is changed OR Exchange has a different number of aliases we are going to re-initialize the aliasses in Zarafa
    if smtp_changed ==1 or ex_smtp_number!=za_smtp_number:	  
        try: 
            if apply== 'y':
                za_ldap.modify_s(dn, [(ldap.MOD_DELETE, za_mail_attribute, None )])
            print "INFO MAIL: Succesfull clear: " + za_mail_attribute + " for user " + dn
        except:
            pass
        for ex_smtp in smtp_ex_value:
            if ex_smtp[0:4] == 'smtp':
                try:
                    if apply== 'y':
                        za_ldap.modify_s(dn, [(ldap.MOD_ADD, za_mail_attribute , ex_smtp[5:])])
                    print "INFO MAIL: Succesfull add: " + za_mail_attribute + " into " +  ex_smtp[5:] + " for user " + dn
                except:
                    print "ERROR MAIL: problem with writing to directory server for " + za_mail_attribute + " into " +  ex_smtp[5:] + " for user " + dn
                    sys.exit()
print "Mail field sync complete"
ex_ldap.unbind()
za_ldap.unbind()
sys.exit()
