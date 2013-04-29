<?php

namespace Tactics\Bundle\ExcelBundle\Helpers;


class ObjectTransformerHelper {
    /**
     * Returns an array of object values based on wanted properties/methods or all properties (providing that a getMethod can be guessed)
     *
     * @param $object
     * @param null $wanted_properties
     * @return array
     */
    public function getWritablePropertiesOfObject($object, $wanted_properties = null)
    {
        $writableProperties = array();
        $reflectionClassObject = new \ReflectionClass(get_class($object));

        if(!$wanted_properties) {
            foreach($reflectionClassObject->getProperties() as $property) {
                $guessedGetter = $this->getGetterForProperty($reflectionClassObject, $property->getName());
                if($guessedGetter) {
                    $writableProperties[$property->getName()] = $object->$guessedGetter();
                }
            }
        }
        else {
            foreach($wanted_properties as $property) {
                $guessedGetter = $this->getGetterForProperty($reflectionClassObject, $property);

                if($guessedGetter) {
                    $writableProperties[$property] = $object->$guessedGetter();
                }
            }
        }

        return $writableProperties;
    }

    /**
     * Guesses the accesor method for a property
     *
     * @param \ReflectionClass $reflectionObject
     * @param $property
     * @return null|string
     */
    protected function getGetterForProperty(\ReflectionClass $reflectionObject, $property)
    {
        $reformedPropertyName = $property;
        $propertyPieces = array();
        foreach (explode('_', $property) as $propertyNamePiece) {
            $propertyPieces[] = ucfirst($propertyNamePiece);
            $reformedPropertyName = implode($propertyPieces);
        }

        $guessedGetter = 'get' . $reformedPropertyName;
        if ($reflectionObject->hasMethod($guessedGetter)) {
            return $guessedGetter;
        }

        return null;
    }
}