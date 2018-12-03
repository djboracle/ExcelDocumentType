/*************************************************************/

   MAKE SURE THE SCHEMA IS WHICH YOU EXECUTE THIS SCRIPT
   HAS PERMISSION TO:

   CREATE SEQUENCES,
   CREATE GLOBAL TEMPORARY TABLES,
   EXECUTE THE SYS.ANYDATA OBJECT,
   EXECUTE THE DBMS_TYPES PACKAGE

   IF YOU ALREADY HAVE SOME OF THE FOLLOWING COMPONENTS INSTALLED
   (i.e. ExcelDocumentType or AnonymousFunction) YOU WILL
   GET THE APPROPRIATE ERROR MESSAGES IN SQLPLUS REMINDING
   YOU THAT CERTAIN OBJECTS ALREADY EXIST.

**************************************************************/


@ExcelDocumentTypeGTT.udt;
@ExcelDocTypeUtils.pkg
