# flowchart-generator

## �T�v
�t���[���L�q�����e�L�X�g�t�@�C������͂Ƃ��āA�G�N�Z���̐}�`�Ńt���[�`���[�g���쐬���܂��B

## �g����
### �e�L�X�g�t�@�C���̏���
�ȉ��̃��[���ŋL�q���܂��B`sample/flowchart.txt`���Q�l�ɂ��Ă��������B

#### String
������͋󔒂��܂ޏꍇ�A`" "`�ł�����܂��B�܂܂Ȃ��ꍇ�͂ǂ���ł����܂��܂���B`"`�̃G�X�P�[�v�ɂ͂܂��Ή����Ă��܂���B
```
aaa
"aaa"
"aaa bbb ccc"
```

#### Statement
- �t���[�`���[�g�J�n
  ```
  # <string>
  ```
- ����
  ```
  <string>
  ```
- ��`�Ϗ���
  ```
  call <string>
  ```
- ��������
  ```
  if <string> [label1,label2]
    {<statement>} | continue
  else
    {<statement>}
  end-if

  if <string> [label1,label2]
    {<statement>}
  end-if
  ```
- �J��Ԃ�
  ```
  do <string> until <string>
    {<statement>}
  loop

  do <string>
    {<statement>}
  loop until <string>
  ```
- �f�[�^
  ```
  read <string> [label]
    {<statement>}
  end-read

  read <string>

  write <string>
  ```
- ����
  ```
  print <string>
  ```
- �\��
  ```
  display <string>
  ```
- ����
  ```
  input <string>
  ```

### �t���[�`���[�g�̍쐬
- `FlowChartGenerator.xlsm`�Ɠ����t�H���_�ɁA�p�ӂ����t�@�C����`flowchart.txt`�Ƃ����t�@�C�����ł����܂��B
- �G�N�Z���t�@�C�����J���A`Generate`�{�^�����N���b�N���܂��B

## ���l
xlsm����̃\�[�X�R�[�h���o�ɂ́A[vbac](https://github.com/vbaidiot/ariawase)���g�p���Ă��܂��B
