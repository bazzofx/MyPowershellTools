$max = 1..10

foreach($object in $max) {
            [pscustomobject]@{
            Incident = $object
            Description = $object
            Group = $object
            'Assigned To' = $object
                  }
}

$object