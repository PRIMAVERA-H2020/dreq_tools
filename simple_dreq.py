"""
A simple interface to the CMIP6 data request.
This was initially written to work with updating spreadsheets taken from the
data request

Written by Matthew Mizielinski (matthew.mizielinski@metoffice.gov.uk)

Crown Copyright (2016)

"""


from dreqPy import dreq
import dreqPy.vrev as vrev
import warnings

dq = None #= dreq.loadDreq()

def initialise(data_request=None, quiet=False):
    """
    Set up the link to the data request.

    Args:
        data_request (object, optional): If the data request is already loaded pass the object
            returned by dreq.loadDreq(). If unset this will be loaded
        quiet (bool, optional): set to True to suppress load message

    """
    global dq

    if data_request is None:
        dq = dreq.loadDreq()
    else:
        dq = data_request

    print "Data request version {0} loaded".format(dq.version)

def lookup_uid(cmor, miptable, first=False, as_object=True):
    """
    interpret the data request and return the unique id (uid)
    for for the supplied CMOR name, miptable combination.

    Args:
        cmor (str): CMOR variable name
        miptable (str): MIP table name
        first (bool, optional): if True and multiple entries are found, return the first*
        as_object (bool, optional): if True return the uid object rather than the string

    Returns:
        uid string (if as_object is False) or data request object corresponding to uid.
        If no variable is found then None is returned

    *would indicate an inconsistency in the data request.
    """

    if dq is None:
        raise Exception("data request not initialised; run initialise() first")

    cmor_uids = dq.inx.CMORvar.label[cmor]
    out_uids = []
    for cid in cmor_uids:
        if dq.inx.uid[cid].mipTable == miptable:
            out_uids.append(cid)

    if len(out_uids) == 0:
        return None
    elif len(out_uids) == 1:
        out_uid = out_uids[0]
        if as_object:
            return dq.inx.uid[out_uid]
        else:
            return out_uid
    else:
        warnings.warn("Found multiple dq.inx.uid for variable %s in table %s:\n" \
                      % (cmor, miptable) + str(out_uids))
        if first and as_object:
            return dq.inx.uid[out_uids[0]]
        elif first and not as_object:
            return out_uids[0]
        else:
            return None


def get_long_name(uid):
    """
    return long name for variable
    """
    return uid.title


def get_units(uid, cf_units=False):
    """
    return units as string, or if cf_units as a cf_units.Unit object
    """
    u = dq.inx.uid[uid.vid].units
    if cf_units:
        from cf_units import Unit
        return Unit(u)
    else:
        return u


def get_description(uid):
    """
    return description for variable
    """
    return dq.inx.uid[uid.vid].description


def get_comment(uid):
    """
    return comment for variable
    """
    return uid.description


def get_var_name(uid):
    """
    return variable name
    """
    return dq.inx.uid[uid.vid].label


def get_standard_name(uid):
    """
    return CF standard name
    """
    return dq.inx.uid[uid.vid].sn


def get_cell_methods(uid):
    """
    return cell_methods as string
    """
    return dq.inx.uid[uid.stid].cell_methods


def get_positive(uid):
    """
    return positive
    """
    return uid.positive


def get_type(uid):
    """
    return variable type
    """
    return uid.type


def get_dimensions(uid):
    """
    return dimensions string from information in variable structures for space and time
    """
    u_stid = dq.inx.uid[uid.stid]
    u_spid = dq.inx.uid[u_stid.spid]
    u_tmid = dq.inx.uid[u_stid.tmid]

    other_dims = u_stid.odims

    if other_dims:
        try:
            other_dims = other_dims.split("|")
        except AttributeError:
            other_dims = []
    else:
        other_dims = []

    retval = " ".join(u_spid.dimensions.split("|") + other_dims + [u_tmid.dimensions])
    if isinstance(u_stid.cids, list):
        for j in u_stid.cids:
            retval += " "+ j.replace("dim:", "")

    return retval


def get_modeling_realm(uid):
    """
    return modeling realm
    """ 
    return uid.modeling_realm


def get_frequency(uid):
    """
    return frequency string
    """
    return uid.frequency


def get_cell_measures(uid):
    """
    return cell measures as a string
    """
    return dq.inx.uid[uid.stid].cell_measures


def get_prov(uid):
    """
    return provenance string
    """
    return uid.prov


def get_provNote(uid):
    """
    return provenance note string
    """
    return uid.provNote


def get_rowIndex(uid):
    """
    return row index as integer
    """ 
    return uid.rowIndex


def get_vid(uid):
    """
    return vid (variable uid) as string
    """
    return uid.vid


def get_uid(uid):
    """
    return uid as string
    """
    return uid.uid


def get_priority(uid):
    """
    return default priority as integer
    """
    return uid.defaultPriority


def get_mips_requesting(uid):
    """
    return string of MIPs requesting variable
    """
    return ",".join(sorted(list(vrev.checkVar(dq).chkCmv(uid.uid))))

if __name__ == "__main__":
    pass
