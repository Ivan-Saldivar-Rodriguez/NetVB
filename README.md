# NetVB
Free code for VB Net development

# WS_Generic_DB
Librería VB Net que expone métodos para automatizar el trabajo con base de datos. Fue desarrollado para proyectos que 
usan SQL Server, pero puede adaptarse fácilmente para otro motores de datos que permitan el uso de procedimientos almacenados.

# LoadBalancerWS
Esta librería permite implementar un balanceo de carga para aplicaciones web, este balanceo se basa en la orquestación de varios
grupos de aplicaciones a nivel del IIS y que apuntan a un mismo código de aplicación web, luego estas múltiples instancias de
la misma aplicación se distirbuyen en forma equitativa las conexiones de usuario a través de esta librería. 
Requiere que se implementen en el servidor de datos utilizado un par de tablas que administran los nombres de las instacias del 
IIS y otra que mantiene las conexiones de usuario. (en el proyecto se incluyen los script SQL para crear esas tablas como los
procedimientos almacenados que se usan para implementar el balanceo).

# DataSerializedTransmission
Esta librería implementa operaciones de serialización de objetos de todo tipo (cadenas de texto, dataset, archivos, etc) y 
deserialización de los mismos.
