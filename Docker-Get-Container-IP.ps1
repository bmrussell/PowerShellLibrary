param ([string] $id)

docker inspect --format '{{ .NetworkSettings.Networks.nat.IPAddress }}' $id
