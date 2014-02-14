"""

Created by: Nathan Starkweather
Created on: 02/14/2014
Created in: PyCharm Community Edition

Module to hold the debug type.
"""


def make_debug_type(batch_type):
    """
    @param batch_type: the type used for pbs lib release builds
    @type batch_type: type
    @return: debugging type
    @rtype: type

    To avoid circular imports, make the class within
    the context of this local function, to be called
    with the argument corresponding to the release type.

    ie from debug_type import make_debug_type
    batchtype = make_debug_type(release_batch_type)
    """

    class PBSBatchDebugType(batch_type):
        """Metaclass for debugging all of PBSlib classes (?)

        Pseudo metaclasses take name, bases, kwargs
        and return them after (possibly) modifying them.
        This makes it easy to make specific behaviors
        more modular, and easier to work with.

        Creating property aliases is a bit tricker, since we
        need to make sure that we don't override inherited
        attributes,
        '''
        """

        from officelib.nsdbg import OverrideWarningMeta, VerboseEmptyMethodMeta, \
                        SlotsNoticeMeta  # , ExplicitVariableDeclarationMeta
        pseudo_meta_list = (
                            OverrideWarningMeta,
                            VerboseEmptyMethodMeta,
                            SlotsNoticeMeta,
    #                             ExplicitVariableDeclarationMeta
                            )

        def __new__(mcs, name, bases, kwargs):

            for pmeta in mcs.pseudo_meta_list:
                name, bases, kwargs = pmeta(name, bases, kwargs)

            new_cls = super().__new__(mcs, name, bases, kwargs)

            return new_cls

    return PBSBatchDebugType
