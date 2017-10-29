#! /bin/bash

#
# Variables de ejecución
PASSWORD_FILE=/tmp/gf.pwd
HOME_GLASSFISH="/servers/glassfish4.1"
DOMAIN_ROOT_DIR="/apps/gf_domains"
NOMBRE_DOMINIO="satelite"
PORT_BASE=9000
ADMIN_USER="admin"
ADMIN_PASS="manager1"
ADMIN_PORT=$[$PORT_BASE + 48]
DB_HOST="192.168.56.11"
DB_PORT=1521
DB_SID="orcl"
DB_USER="satelite"
DB_PASS="satelite"
AS_INSTALL=$HOME_GLASSFISH/glassfish
AS_ADMIN="$AS_INSTALL/bin/asadmin"
AS_ADMIN_AUTH="$AS_ADMIN --user $ADMIN_USER --passwordfile $PASSWORD_FILE --port $ADMIN_PORT"

#
# Termina la ejecución del script con valor de retorno error
function terminarEjecucion 
{
   deletePasswordFile;
   exit 1;
}

#
# Termina la ejecución del script en con retorno exitoso
function salir 
{
   deletePasswordFile;
   exit 0;
}

#
# Termina la ejecución del script siembre y cuando no haya errores
function checkEjecucion 
{
   if [ "$?" != "0" ]; then
      deletePasswordFile;
      exit $?;
   fi;   
}

#
# Lectura e inicialización de variables
function showParameters
{
   echo ""
   echo "VALORES DE EJECUCION"
   echo "===================="
   echo "AS_INSTALL    :" $AS_INSTALL
   echo "AS_ADMIN      :" $AS_ADMIN
   echo "DOMINIO       :" $NOMBRE_DOMINIO
   echo "PORT BASE     :" $PORT_BASE
   echo "ADMIN PORT    :" $ADMIN_PORT
   echo "ADMIN USER    :" $ADMIN_USER
   echo "ADMIN PASSWORD:" $ADMIN_PASS
   echo "DOMAIN ROOT   :" $DOMAIN_ROOT_DIR
   echo "DB.HOST       :" $DB_HOST
   echo "DB.PORT       :" $DB_PORT
   echo "DB.SID        :" $DB_SID
   echo "DB.USERNAME   :" $DB_USER
   echo "DB.PASSWORD   :" $DB_PASS
   echo ""
}

#
# Función que crea el archivo con la password del usuario "admin"
function createPasswordFile 
{
   echo "AS_ADMIN_PASSWORD="$ADMIN_PASS > $PASSWORD_FILE
   checkEjecucion;
}

#
# Elimina el archivo que contiene la password del usuario "admin"
function deletePasswordFile 
{
   rm -f $PASSWORD_FILE
}

#
# Función que inicia la ejecución de un Dominio
function startDomain 
{
   local CHECK_DOMINIO=`$AS_ADMIN list-domains --domaindir $DOMAIN_ROOT_DIR|grep "^$NOMBRE_DOMINIO\ "`
   
   if [ "$CHECK_DOMINIO" == "" ]; then
      echo "El Dominio $NOMBRE_DOMINIO no existe en el servidor";
      terminarEjecucion;
   fi;
   
   if [ "`echo $CHECK_DOMINIO|grep \"not running\"`" != "" ]; then
      echo ""
      echo "Iniciando Dominio" $NOMBRE_DOMINIO "..."
      $AS_ADMIN start-domain --domaindir $DOMAIN_ROOT_DIR $NOMBRE_DOMINIO
      checkEjecucion;
   fi;
}

#
# Functión que detiene la ejecución de un Dominio
function stopDomain 
{
   local CHECK_DOMINIO=`$AS_ADMIN list-domains --domaindir $DOMAIN_ROOT_DIR|grep "^$NOMBRE_DOMINIO\ "`
   
   if [ -z "$CHECK_DOMINIO" ]; then
      echo "El Dominio $NOMBRE_DOMINIO no existe en el servidor";
      terminarEjecucion;
   fi;
   
   #
   # Si el dominio no está en ejecución, entonces nada más que hacer
   if [ "`echo $CHECK_DOMINIO|grep \"not running\"`" != "" ]; then
      return;
   fi;
   
   echo ""
   echo "Deteniendo Dominio" $NOMBRE_DOMINIO "..."
   $AS_ADMIN stop-domain --domaindir $DOMAIN_ROOT_DIR $NOMBRE_DOMINIO
   checkEjecucion;
}

#
# Función para reiniciar el dominio
function restartDomain
{
    echo ""
    echo "Reiniciando el dominio:" $NOMBRE_DOMINIO "..."
    $AS_ADMIN restart-domain --domaindir $DOMAIN_ROOT_DIR $NOMBRE_DOMINIO
    checkEjecucion;
}

#
# Función que crea el Dominio
function createDomain 
{
    #
    # Nos aseguramos que exista el directorio de destino
    if ! [ -d $DOMAIN_ROOT_DIR ]; then
      mkdir -p $DOMAIN_ROOT_DIR
      terminarEjecucion;
    fi;

    local CHECK_DOMINIO=`$AS_ADMIN list-domains --domaindir $DOMAIN_ROOT_DIR|grep "^$NOMBRE_DOMINIO\ "`

    #
    # Primero se valida que el dominio no exista
    if [ "$CHECK_DOMINIO" != "" ]; then
      echo "El Dominio $NOMBRE_DOMINIO ya existe en este servidor";
      terminarEjecucion;
    fi;

    #
    # Creación de archivo con la password
    createPasswordFile;

    #
    # Creación del Dominio
    $AS_ADMIN --user $ADMIN_USER --passwordfile $PASSWORD_FILE \
             create-domain --savemasterpassword=false \
                           --domaindir $DOMAIN_ROOT_DIR \
                           --portbase $PORT_BASE \
                           $NOMBRE_DOMINIO
    checkEjecucion;

    #
    # Se habilita la administración por Consola de manera segura
    echo ""
    echo "Habilitando administración segura..."   
    $AS_ADMIN start-domain --domaindir $DOMAIN_ROOT_DIR $NOMBRE_DOMINIO
    checkEjecucion;

    #
    # Para habilitar la administración segura es preciso
    # autenticarse en la consola del dominio
    $AS_ADMIN_AUTH enable-secure-admin
    checkEjecucion;

    #
    # Reinicio del dominio
    restartDomain;

    #
    # Se eliminan valores por defecto del Dominio
    echo ""
    echo "Eliminación de valores por defecto de Dominio:" $NOMBRE_DOMINIO "..."
    for jdbc in `$AS_ADMIN_AUTH list-jdbc-resources server|grep -va successfully`; do
      $AS_ADMIN_AUTH delete-jdbc-resource $jdbc;
      checkEjecucion;
    done;

    for pool in `$AS_ADMIN_AUTH list-jdbc-connection-pools|grep -va successfully|grep -va target`; do
      $AS_ADMIN_AUTH delete-jdbc-connection-pool $pool;
      checkEjecucion;
    done;
}

#
# Función que crea un JDBC - ORACLE
function createJdbcOracle
{
   local NAME=$1
   local DB_HOST=$2
   local DB_PORT=$3
   local DB_SID=$4
   local DB_USER=$5
   local DB_PASS=$6
   
   local POOL_NAME="$NAME""Pool"
   local JNDI_NAME="jdbc/$NAME"
   local JDBC_URL="jdbc\:oracle\:thin\:@$DB_HOST\:$DB_PORT\:$DB_SID"
   local CHECK_POOL=`$AS_ADMIN_AUTH list-jdbc-connection-pools|grep $POOL_NAME`
      
   #
   # Eliminamos la configuración actual
   if [ "$CHECK_POOL" != "" ]; then
      $AS_ADMIN_AUTH delete-jdbc-connection-pool --cascade true $POOL_NAME
      checkEjecucion;
   fi;
   
   #
   # Creación del Pool JDBC
   $AS_ADMIN_AUTH create-jdbc-connection-pool \
                  --restype javax.sql.ConnectionPoolDataSource \
                  --datasourceclassname oracle.jdbc.pool.OracleConnectionPoolDataSource \
                  --ping true \
                  --steadypoolsize 8 \
                  --maxpoolsize 32 \
                  --property URL="$JDBC_URL":user=$DB_USER:password=$DB_PASS \
                  $POOL_NAME
   checkEjecucion;

   #
   # Creación del DataSource
   $AS_ADMIN_AUTH create-jdbc-resource \
                  --enabled true \
                  --connectionpoolid $POOL_NAME \
                  $JNDI_NAME
}

#
# Función que permite eliminar una opción desde la JVM
function deleteJvmOption 
{
   local JVM_OPTION=$1
   local CHECK_JVM_OPTION=`$AS_ADMIN_AUTH list-jvm-options|grep "\-$JVM_OPTION"`

   if [ "$CHECK_JVM_OPTION" != "" ]; then
      echo "Eliminando: -$JVM_OPTION"
      $AS_ADMIN_AUTH delete-jvm-options "-$JVM_OPTION"
      checkEjecucion;
   fi;
}

#
# Función que permite agregar una opción en la JVM
function createJvmOption 
{
   local JVM_OPTION=$1
   local CHECK_JVM_OPTION=`$AS_ADMIN_AUTH list-jvm-options|grep "\-$JVM_OPTION"`

   echo "Creando: -$JVM_OPTION"
   if [ "$CHECK_JVM_OPTION" != "" ]; then
      return;
   fi;
   
   $AS_ADMIN_AUTH create-jvm-options "-$JVM_OPTION"
   checkEjecucion;
}

#
# Función que configura los parámetros de ejecución de la JVM
function setupJVM 
{
   deleteJvmOption "client"
   deleteJvmOption "Xmx512m"
   deleteJvmOption "XX\:MaxPermSize=192m"
   createJvmOption "server"
   createJvmOption "Xmx1024m"
   createJvmOption "Xms1024m"
   createJvmOption "verbose\:gc"
   createJvmOption "XX\:MaxPermSize=256m"
   createJvmOption "XX\:+PrintGCDateStamps"
   createJvmOption "XX\:+PrintGCDetails"
   createJvmOption "XX\:-HeapDumpOnOutOfMemoryError"
   createJvmOption "XX\:HeapDumpPath=\${com.sun.aas.instanceRoot}/logs/glassfish.hprof"
   createJvmOption "Xloggc\:\${com.sun.aas.instanceRoot}/logs/gc.log"
}

#
# Configura el Dominio
function setupDomain 
{
   #
   # Nos aseguramos que el Dominio esté iniciado
   startDomain;

   #
   # Nos aseguramos que exista el archivo con la password del ADMIN
   createPasswordFile;

   #
   # Nos aseguramos de copiar los JARs
   cp -f jdbc/*.jar $DOMAIN_ROOT_DIR/$NOMBRE_DOMINIO/lib/ext
   cp -f libs/*.jar $DOMAIN_ROOT_DIR/$NOMBRE_DOMINIO/lib/
   
   #
   # Reinicio del dominio, para que carguen los drivers JDBCs
   restartDomain;

   #
   # Configuración del Pool a la Base Datos
   echo ""
   echo "Configurando Data Source JDBC..."
   createJdbcOracle "satelite" "$DB_HOST" "$DB_PORT" "$DB_SID" "$DB_USER" "$DB_PASS"

   #
   # Configuración de la VM
   echo ""
   echo "Configurando parámetros de la JVM..."
   setupJVM;

   #
   # Configuración de los Thread Pools
   echo ""
   echo "Configurando de los Thread Pools..."
   $AS_ADMIN_AUTH set server.thread-pools.thread-pool.http-thread-pool.max-queue-size=4096
   $AS_ADMIN_AUTH set server.thread-pools.thread-pool.http-thread-pool.max-thread-pool-size=200
   $AS_ADMIN_AUTH set server.thread-pools.thread-pool.http-thread-pool.min-thread-pool-size=200
   
   #
   # Finalmente se reinicia el Dominio para que se apliquen las configuraciones
   echo ""
   echo "Reiniciando el dominio:" $NOMBRE_DOMINIO "..."
   $AS_ADMIN restart-domain --domaindir $DOMAIN_ROOT_DIR $NOMBRE_DOMINIO
   checkEjecucion;
}

#
# Se leen por pantalla las variables de ambiente
showParameters;
createDomain;
setupDomain;
salir;
