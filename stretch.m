function output=stretch(p,s)
id=find(p=='-');

output=s;
count=0;
if ~isempty(id)
 for i=1:length(id)
    position=(id(i)-1+count)*3;
    output = insertAfter(output,position,'---');
    count=count+1;
 end
end