<?php
libxml_disable_entity_loader(false);

class SoftariusOOXML extends ZipArchive
{
    public $tmddir;
    public $contentType = array();
    public $customProps;
    public $OnProp;

    public function relatedFilename($zipentry)
    {
        return substr($zipentry, 1);
    }

    public function getCustomProps()
    {
        if (!$this->customProps) {
            return null;
        }
        $props = $this->customProps->getElementsByTagName('property');
        if ($this->OnProp) {
            foreach ($props as $prop) {
                $name = $prop->getAttribute('name');
                $valnodes = $prop->getElementsByTagNameNS('http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes', '*');
                $vn = $valnodes[0];
                $prop = new StdClass();
                $prop->key = $name;
                $prop->val = $vn->nodeValue;
                $prop->kind = $vn->localName;

                call_user_func($this->OnProp, $prop);
            }
        }

        return $props;
    }

    public function setSettings($settings)
    {
        //echo $this->tmpdir;
        $sfn = 'word/settings.xml';
        $tfn = $this->tmpdir . '/' . $sfn;
        $this->extractTo($this->tmpdir, $sfn);
        $this->settings = new DomDocument();
        $this->settings->load($tfn);
        $updateFields = $this->settings->createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:updateFields');
        $updateFields->setAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:val', 'true');
        $this->settings->documentElement->appendChild($updateFields);
        $this->settings->save($tfn);
        $this->addFile($tfn, $sfn);
    }

    public function open($filename, $flags = null)
    {
        $r = parent::open($filename, $flags);
        $this->tmpdir = tempdir();

        $r = $this->extractTo($this->tmpdir);

        $ct = new DomDocument();
        $ct->load($this->tmpdir . '\[Content_Types].xml');


        $or = $ct->getElementsByTagName('Override');
        foreach ($or as $o) {
            if ($o->getAttribute('ContentType') == 'application/vnd.openxmlformats-officedocument.custom-properties+xml') {
                $fn = $this->relatedFilename($o->getAttribute('PartName'));
                $r = $this->extractTo($this->tmpdir, $fn);
                $this->customProps = new DomDocument();
                $this->customProps->load($this->tmpdir . $o->getAttribute('PartName'));
            }
        }
    }

    public function getCustomSQL()
    {
        $dom = new DomDocument();
        $dom->load($this->tmpdir . '\customXml\item1.xml');

        $root = $dom->documentElement;
       
        $schema = $root->localName;
        $r = '';
        $r .= "drop schema if exists \"$schema\" cascade;\n";
        $r .= "create schema if not exists \"$schema\" ;\n";
        $r .= "comment on schema \"$schema\" is '{$root->namespaceURI}'; \n";



        foreach ($root->childNodes as $table) {
            if ($table->localName) {
                $tn = "\"$schema\".\"{$table->localName}\"";
                $r .= "\n create table $tn ( \n";
                $fields = [];
                $v = null;
                $values = null;
                foreach ($table->childNodes as $rec) {
                    if ($rec->localName) {
                        $rowname = $rec->localName;
                        foreach ($rec->attributes as $attr) {
                            $fields[$attr->name] = '';
                            $values[$attr->name] = $attr->value;
                        }
                        $v[] = $values;
                    }
                }

                $f = '"' . implode('" varchar(250), "', array_keys($fields)) . '" varchar(250) ) inherits (ds.tabledata);' . "\n";
                $r .= $f;

                $r .= "comment on table $tn is '{$rowname}';\n";
               /* foreach ($v as $val) {
                    $keys = array_keys($val);
                    $k = '"' . implode('", "', $keys) . '"';
                    $values = "'" . implode("', '", $val) . "'";
                    $r .= "insert into $tn ($k) values ($values); \n";
                }*/
            }
        }
        return $r;
    }
}

/**
 * Making a temporary dir for unpacking a zipfile into
 *
 * @link https://stackoverflow.com/questions/1707801/making-a-temporary-dir-for-unpacking-a-zipfile-into
 *
 * @return string
 */
function tempdir()
{
    $tempfile = tempnam(sys_get_temp_dir(), '');
    // you might want to reconsider this line when using this snippet.
    // it "could" clash with an existing directory and this line will
    // try to delete the existing one. Handle with caution.
    if (file_exists($tempfile)) {
        unlink($tempfile);
    }
    mkdir($tempfile);
    if (is_dir($tempfile)) {
        return $tempfile;
    }
}