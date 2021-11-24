function [d,s,s0]=p2d(p,codon)
for i=1:length(p)
    c=p(i);
   id=strmatch(c, char(codon(:,1)));   %look up AA{3} in condon table 1
   if ~isempty(id)
    d{i}=codon{id,2};  %find the corresponding gene in table
   else
    d{i}='---';
   end
end

s0=sprintf('%s',d{:});
id=find(s0~='-');
s=s0(id);