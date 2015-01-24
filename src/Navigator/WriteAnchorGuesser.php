<?php
/*
* The MIT License (MIT)
*
* Copyright (c) 2014 Artur Augustyniak
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files (the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/

namespace Aaugustyniak\PhpExcelHandler\Navigator;


/**
 * @author Artur Augustyniak <artur@aaugustyniak.pl>
 */
class WriteAnchorGuesser
{

    /**
     * @var array
     */
    private $payload;

    /**
     * @var CellNavigator
     */
    private $navigator;

    /**
     *
     * @param array $payload
     */
    function __construct(array $payload = null)
    {
        if (null !== $payload) {
            $this->setPayload($payload);
        }
        $this->navigator = new CellNavigator();
    }

    /**
     * @return array
     */
    public function getPayload()
    {
        return $this->payload;
    }

    /**
     * @return CellNavigator
     */
    public function getNavigator()
    {
        return $this->navigator;
    }


    /**
     * @param CellNavigator $navigator
     */
    public function setNavigator(CellNavigator $navigator)
    {
        $this->navigator = $navigator;
    }


    /**
     * Change payload
     * @param array $payload
     */
    public function setPayload(array $payload)
    {
        $this->payload = $this->fixIndexing($payload);
    }

    /**
     * Make all indices numeric
     */
    public function forceFixIndexing()
    {
        $this->payload = \array_values($this->payload);
    }

    /**
     * non numeric indices can result as double zero index
     * @see http://stackoverflow.com/a/3260454
     * @param array $input
     * @return array
     */
    private function fixIndexing(array $input)
    {
        $keys = array_keys($input);
        if (!empty($keys)) {
            $lastKey = $keys[0];
            for ($i = 1; $i < count($keys); $i++) {
                if ($keys[$i] == $lastKey) {
                    return \array_values($input);
                }
                $lastKey = $keys[$i];
            }
        }
        return $input;

    }


    /**
     * Guess top left corner of write area
     * guess is based on 2D array indices
     *
     * @return string PHPExcel cell address
     */
    public function getWriteAnchor()
    {
        return $this->navigator->getAddressFor(
            $this->getColumnFromPayload(),
            $this->getRowFromPayload());
    }


    /**
     * Get write start column index
     * @return int
     */
    public function getColumnFromPayload()
    {
        $keys = array_keys($this->payload);
        if (!empty($keys)) {
            return (int)$keys[0];
        } else {
            return 0;
        }
    }

    /**
     * Get write start row index
     * @return int
     */
    public function getRowFromPayload()
    {
        if (!empty($this->payload)) {
            $column = $this->getColumnFromPayload();
            if (is_array($this->payload[$column])) {
                $keys = array_keys($this->payload[$column]);
                return (int)$keys[0];
            } else {
                return 0;
            }
        } else {
            return 0;
        }
    }

}